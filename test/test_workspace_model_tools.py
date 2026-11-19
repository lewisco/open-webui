"""Run with pytest or directly with the backend's Python interpreter."""

import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory


def test_workspace_model_tools():
    # The backend configures its database and runs migrations at import time.
    with TemporaryDirectory() as root:
        env = {
            **{key: os.environ[key] for key in ('PATH', 'HOME', 'SYSTEMROOT') if key in os.environ},
            'PYTHONPATH': str(Path(__file__).resolve().parents[1] / 'backend'),
            'DATA_DIR': root,
            'STATIC_DIR': f'{root}/static',
            'DATABASE_URL': f'sqlite:///{root}/test.db',
            'USE_SLIM_DOCKER': 'true',
            'OFFLINE_MODE': 'true',
            'WEBUI_SECRET_KEY': 'workspace-model-tools-test',
        }
        subprocess.run([sys.executable, __file__, '--isolated'], env=env, check=True)


async def exercise_model_tools():
    from fastapi import Request
    from open_webui.main import app
    from open_webui.models.access_grants import AccessGrants
    from open_webui.models.config import Config
    from open_webui.models.files import FileForm, Files
    from open_webui.models.functions import FunctionForm, Functions
    from open_webui.models.groups import GroupForm, Groups
    from open_webui.models.knowledge import KnowledgeForm, Knowledges
    from open_webui.models.model_history import ModelHistories
    from open_webui.models.models import ModelForm, Models
    from open_webui.models.notes import NoteForm, Notes
    from open_webui.models.skills import SkillForm, Skills
    from open_webui.models.tools import ToolForm, Tools
    from open_webui.models.users import Users
    from open_webui.utils.middleware import execute_tool_call
    from open_webui.utils.tools import get_builtin_tools

    admin, owner, reader, outsider = [
        await Users.insert_new_user(id, id, f'{id}@example.org', role=role)
        for id, role in [('admin', 'admin'), ('owner', 'user'), ('reader', 'user'), ('outsider', 'user')]
    ]
    await Config.upsert({'user.permissions': {'workspace': {'models': False}}})
    authors = await Groups.insert_new_group(
        admin.id, GroupForm(name='Authors', description='', permissions={'workspace': {'models': True}})
    )
    await Groups.add_users_to_group(authors.id, [owner.id, reader.id])
    request = Request({'type': 'http', 'app': app, 'headers': [], 'path': '/api/chat/completions'})
    app.state.MODELS = {id: {'id': id, 'owned_by': 'openai'} for id in ('base', 'private-base')}
    for id, user in [('base', owner), ('private-base', admin)]:
        await Models.insert_new_model(ModelForm(id=id, name=id, params={}, meta={}), user.id)

    async def registry(user, enabled=None, current='authoring-model'):
        return await get_builtin_tools(
            request,
            {'__user__': user.model_dump()},
            model={'id': current, 'info': {'meta': {'builtinTools': {'models': enabled}}}}
            if enabled is not None
            else {'id': current},
        )

    async def call(tools, tool, **args):
        _, result, _, _ = await execute_tool_call(
            {}, {'tools': tools}, None, {'name': tool, 'arguments': json.dumps(args)}
        )
        return json.loads(result) if isinstance(result, str) and result.startswith('{') else result

    names = {'create_model', 'update_model', 'list_models', 'view_model'}
    assert names.isdisjoint(await registry(admin, enabled=False))
    assert names.isdisjoint(await registry(outsider))
    assert names <= (await registry(admin)).keys()
    request.state.internal = True
    assert names.intersection(await registry(owner)) == {'list_models', 'view_model'}
    request.state.internal = False
    tools = await registry(owner)
    assert names <= tools.keys()
    for name in names:
        assert not any(k.startswith('__') for k in tools[name]['spec']['parameters']['properties'])

    await Tools.insert_new_tool(owner.id, ToolForm(id='sample', name='Sample', content='', meta={}), [])
    await Skills.insert_new_skill(owner.id, SkillForm(id='skill', name='Skill', content='Instructions'))
    kb = await Knowledges.insert_new_knowledge(owner.id, KnowledgeForm(name='Manual', description=''))
    file = await Files.insert_new_file(owner.id, FileForm(id='manual', filename='manual.txt', path='manual.txt'))
    note = await Notes.insert_new_note(owner.id, NoteForm(title='Intake'))
    for kind in ('filter', 'action'):
        await Functions.insert_new_function(admin.id, kind, FunctionForm(id=kind, name=kind, content='', meta={}))
        await Functions.update_function_by_id(kind, {'is_active': True}, user_id=admin.id)
    model_id = 'support/a&b'
    args = {'id': model_id, 'name': 'Support', 'base_model_id': 'base'}
    meta = {
        'tags': ['ops'],
        'toolIds': ['sample'],
        'skillIds': ['skill'],
        'capabilities': {'vision': False},
        'knowledge': [
            {'type': 'collection', 'id': kb.id},
            {'type': 'file', 'id': file.id},
            {'type': 'note', 'id': note.id},
        ],
        'filterIds': ['filter'],
        'actionIds': ['action'],
        'suggestion_prompts': [{'title': ['Intake', 'Open a ticket'], 'content': 'Start intake'}],
    }
    saved = await call(
        tools, 'create_model', **args, params={'system': 'Ask for a ticket', 'temperature': 0.2}, meta=meta
    )
    assert saved == {'id': model_id, 'url': '/workspace/models/edit?id=support%2Fa%26b'}, saved
    model = await Models.get_model_by_id(model_id)
    assert model.access_grants == [] and model.is_active
    assert [item['name'] for item in model.meta.knowledge] == ['Manual', 'manual.txt', 'Intake']
    assert model.meta.filterIds == ['filter'] and model.meta.actionIds == ['action']
    assert (await call(tools, 'list_models', tag='ops', query='Support'))['total'] == 1
    assert 'error' in await call(tools, 'create_model', **args)
    assert 'error' in await call(tools, 'create_model', id='other', name='Other', base_model_id='private-base')
    assert 'error' in await call(await registry(owner, current=model_id), 'update_model', id=model_id, name='Changed')
    assert 'error' in await call(tools, 'update_model', id='base', name='Changed')

    viewer = await registry(reader)
    assert (await call(viewer, 'list_models'))['total'] == 0
    assert 'error' in await call(viewer, 'view_model', id=model_id)
    grant = {'principal_type': 'user', 'principal_id': reader.id, 'permission': 'read'}
    await AccessGrants.set_access_grants('model', model_id, [grant])
    assert (await call(viewer, 'view_model', id=model_id))['params'] == {}
    assert 'error' in await call(viewer, 'update_model', id=model_id, name='Changed')
    await AccessGrants.set_access_grants('model', model_id, [grant, {**grant, 'permission': 'write'}])
    for key in ('toolIds', 'skillIds', 'knowledge'):
        denied = await call(viewer, 'update_model', id=model_id, meta={key: meta[key]})
        assert 'error' in denied, denied
    for key in ('actionIds', 'filterIds'):
        assert 'error' in await call(tools, 'update_model', id=model_id, meta={key: ['missing']})
    assert len(await ModelHistories.get_history_by_model_id(model_id)) == 1
    await Models.toggle_model_by_id(model_id)
    assert (
        await call(
            viewer,
            'update_model',
            id=model_id,
            params={'system': 'Ask for the ticket number first'},
            meta={'capabilities': {'memory': False}},
            commit_message='Clarify intake',
        )
        == saved
    )
    changed = await Models.get_model_by_id(model_id)
    assert not changed.is_active
    assert changed.params.temperature == 0.2 and changed.params.system == 'Ask for the ticket number first'
    assert changed.meta.capabilities == {'vision': False, 'memory': False}
    assert changed.meta.knowledge == model.meta.knowledge and changed.meta.toolIds == ['sample']
    assert {g.permission for g in changed.access_grants} == {'read', 'write'} and changed.user_id == owner.id
    stale = await call(tools, 'update_model', id=model_id, name='Stale edit', version_id=model.version_id)
    assert stale == {'error': {'code': 'version_conflict'}}
    assert (await Models.get_model_by_id(model_id)).version_id == changed.version_id
    history = await ModelHistories.get_history_by_model_id(model_id)
    assert len(history) == 2 and any(h.commit_message == 'Clarify intake' and h.user_id == reader.id for h in history)


if __name__ == '__main__':
    if sys.argv[1:] == ['--isolated']:
        asyncio.run(exercise_model_tools())
    else:
        test_workspace_model_tools()
