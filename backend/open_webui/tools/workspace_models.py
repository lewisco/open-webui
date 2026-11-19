"""Chat tools for the workspace model editor."""

from typing import Literal
from urllib.parse import urlencode

from fastapi import Request
from pydantic import BaseModel, ConfigDict, Field

from open_webui.config import BYPASS_ADMIN_ACCESS_CONTROL
from open_webui.env import BYPASS_MODEL_ACCESS_CONTROL, ENABLE_FUNCTIONS
from open_webui.models.config import Config
from open_webui.models.functions import Functions
from open_webui.models.models import ModelForm, PromptSuggestion
from open_webui.models.users import UserModel
from open_webui.utils.access_control import has_permission
from open_webui.utils.json_codec import JSONCodec


class KnowledgeReference(BaseModel):
    model_config = ConfigDict(extra='forbid')
    type: Literal['collection', 'file', 'note']
    id: str
    context: Literal['full'] | None = None


class ModelMetadataPatch(BaseModel):
    model_config = ConfigDict(extra='forbid')
    description: str | None = None
    tags: list[str] | None = None
    suggestion_prompts: list[PromptSuggestion] | None = None
    toolIds: list[str] | None = None
    skillIds: list[str] | None = None
    knowledge: list[KnowledgeReference] | None = None
    filterIds: list[str] | None = None
    defaultFilterIds: list[str] | None = None
    actionIds: list[str] | None = None
    builtinTools: dict[str, bool] | None = None
    capabilities: dict[str, bool] | None = None
    defaultFeatureIds: list[str] | None = None


class ModelParameterPatch(BaseModel):
    model_config = ConfigDict(extra='forbid', allow_inf_nan=False)
    system: str | None = None
    temperature: float | None = None
    top_p: float | None = None
    top_k: int | None = None
    max_tokens: int | None = None
    num_ctx: int | None = Field(default=None, gt=0, description='Ollama context window in tokens.')
    seed: int | None = None
    stop: list[str] | None = None
    frequency_penalty: float | None = None
    presence_penalty: float | None = None
    reasoning_effort: str | None = None


async def _knowledge_name(item, request, user):
    from open_webui.routers.files import get_file_by_id
    from open_webui.routers.knowledge import get_knowledge_by_id
    from open_webui.routers.notes import get_note_by_id

    if item['type'] == 'file':
        return (await get_file_by_id(item['id'], user, None)).filename
    if item['type'] == 'note':
        return (await get_note_by_id(request, item['id'], user, None)).title
    return (await get_knowledge_by_id(item['id'], user, None)).name


async def _validated_meta(meta, request, user, existing=None):
    from open_webui.routers.skills import authorized_skill
    from open_webui.routers.tools import get_tools

    patch = ModelMetadataPatch.model_validate(meta or {}).model_dump(exclude_none=True)
    if patch.get('toolIds'):
        tools = await get_tools(request, user=user, db=None)
        if set(patch['toolIds']).difference(tool.id for tool in tools):
            raise ValueError('Tool unavailable')
    for skill_id in patch.get('skillIds') or []:
        if not (await authorized_skill(skill_id, user)).is_active:
            raise ValueError('Skill unavailable')
    for key, kind in [('filterIds', 'filter'), ('actionIds', 'action')]:
        if patch.get(key):
            available = await Functions.get_active_function_ids_by_type(kind) if ENABLE_FUNCTIONS else []
            if set(patch[key]).difference(id for id, _ in available):
                raise ValueError(f'{kind.capitalize()} unavailable')
    if 'defaultFilterIds' in patch or 'filterIds' in patch:
        selected = patch.get('filterIds', (existing or {}).get('filterIds')) or []
        if 'defaultFilterIds' in patch and set(patch['defaultFilterIds']).difference(selected):
            raise ValueError('Default filters must be selected filters')
        if 'filterIds' in patch and 'defaultFilterIds' not in patch and existing and existing.get('defaultFilterIds'):
            patch['defaultFilterIds'] = [id for id in existing['defaultFilterIds'] if id in selected]
    for item in patch.get('knowledge') or []:
        item['name'] = await _knowledge_name(item, request, user)
    return patch


async def _check_base(request, user, base_model_id):
    from open_webui.utils.models import check_model_access, get_all_models

    if not request.app.state.MODELS:
        await get_all_models(request, user=user)
    base = request.app.state.MODELS.get(base_model_id)
    if (
        not base
        or base.get('preset')
        or base.get('owner_id')
        or base.get('owned_by') in ('terminal', 'arena')
        or (base.get('pipeline') or {}).get('type') == 'filter'
    ):
        raise ValueError('Select an available base model from the model editor')
    if not BYPASS_MODEL_ACCESS_CONTROL and not (user.role == 'admin' and BYPASS_ADMIN_ACCESS_CONTROL):
        await check_model_access(user, base)
    if (
        (base.get('info') or {}).get('meta', {}).get('hidden')
        and user.role != 'admin'
        and not await has_permission(user.id, 'workspace.models_hidden', await Config.get('user.permissions'))
    ):
        raise ValueError('Base model unavailable')


def _model_versions(metadata):
    if metadata is None:
        return {}
    context = metadata.get('chat_context') or {}
    metadata['chat_context'] = context
    return context.setdefault('model_versions', {})


def _saved_model(model, metadata):
    _model_versions(metadata)[model['id']] = model['version_id']
    return JSONCodec.dumps({'id': model['id'], 'url': '/workspace/models/edit?' + urlencode({'id': model['id']})})


async def list_models(query: str = None, tag: str = None, page: int = 1, __user__: dict = None) -> str:
    """List readable workspace models, optionally filtered by text or tag. Pages start at 1."""
    from open_webui.routers.models import get_models

    try:
        result = await get_models(query=query, tag=tag, page=page, user=UserModel(**__user__), db=None)
        return JSONCodec.dumps(
            {
                'items': [
                    {
                        **m.model_dump(include={'id', 'name', 'base_model_id', 'is_active'}),
                        'description': m.meta.description,
                    }
                    for m in result.items
                ],
                'total': result.total,
            }
        )
    except Exception as error:
        return JSONCodec.dumps({'error': getattr(error, 'detail', str(error))})


async def view_model(id: str, __user__: dict = None, __metadata__: dict = None) -> str:
    """View a workspace model's configuration. Inference parameters and the system prompt require write access."""
    from open_webui.routers.models import get_model_by_id

    try:
        model = await get_model_by_id(id, UserModel(**__user__), None)
        if not model.base_model_id:
            raise ValueError('Select a workspace model')
        _model_versions(__metadata__)[id] = model.version_id
        return JSONCodec.dumps(
            model.model_dump(
                include={'id', 'name', 'base_model_id', 'params', 'meta', 'is_active', 'write_access', 'version_id'}
            )
        )
    except Exception as error:
        return JSONCodec.dumps({'error': getattr(error, 'detail', str(error))})


async def create_model(
    id: str,
    name: str,
    base_model_id: str = None,
    params: ModelParameterPatch = None,
    meta: ModelMetadataPatch = None,
    __request__: Request = None,
    __user__: dict = None,
    __model_id__: str = None,
    __metadata__: dict = None,
) -> str:
    """Create a private workspace model based on an available provider model.

    :param id: New unique model ID.
    :param base_model_id: Provider model ID. Defaults to the current chat model’s base.
    :param params: Inference parameters, including system for the system prompt.
    :param meta: Description, settings and accessible workspace attachments.
    """
    from open_webui.routers.models import create_new_model

    try:
        user = UserModel(**__user__)
        if base_model_id is None:
            current = __request__.app.state.MODELS.get(__model_id__) or {}
            base_model_id = (current.get('info') or {}).get('base_model_id') or __model_id__
        await _check_base(__request__, user, base_model_id)
        if id in __request__.app.state.MODELS:
            raise ValueError('Model ID already in use')
        form = ModelForm(
            id=id,
            name=name,
            base_model_id=base_model_id,
            access_grants=[],
            params=ModelParameterPatch.model_validate(params or {}).model_dump(exclude_unset=True),
            meta=await _validated_meta(meta, __request__, user),
        )
        return _saved_model(await create_new_model(__request__, form, user, None), __metadata__)
    except Exception as error:
        return JSONCodec.dumps({'error': getattr(error, 'detail', str(error))})


async def update_model(
    id: str,
    name: str = None,
    base_model_id: str = None,
    params: ModelParameterPatch = None,
    meta: ModelMetadataPatch = None,
    commit_message: str = None,
    version_id: str = None,
    __request__: Request = None,
    __user__: dict = None,
    __model_id__: str = None,
    __metadata__: dict = None,
) -> str:
    """Update selected fields on a writable workspace model and save a version. Lists replace the corresponding list.

    :param params: Changed inference parameters, including system for the system prompt.
    :param meta: Changed metadata fields. Supply the full desired list when changing attachments.
    :param commit_message: Optional description for the version history.
    :param version_id: Pass the version_id from view_model to reject stale edits.
    """
    from open_webui.routers.models import authorized_model_history, update_model_by_id

    try:
        user = UserModel(**__user__)
        if id == __model_id__:
            raise ValueError('Choose a different model from the one answering this chat')
        if user.role != 'admin' and not await has_permission(
            user.id, 'workspace.models', await Config.get('user.permissions')
        ):
            raise ValueError('Model authoring permission required')
        model = await authorized_model_history(id, user, None)
        if not model.base_model_id:
            raise ValueError('Select a workspace model')
        if base_model_id is not None:
            if base_model_id == id:
                raise ValueError('A model cannot be its own base')
            await _check_base(__request__, user, base_model_id)
        metadata = model.meta.model_dump()
        changes = await _validated_meta(meta, __request__, user, metadata)
        for key in ('builtinTools', 'capabilities'):
            if isinstance(changes.get(key), dict):
                changes[key] = {**(metadata.get(key) or {}), **changes[key]}
        form = ModelForm(
            id=id,
            name=model.name if name is None else name,
            base_model_id=model.base_model_id if base_model_id is None else base_model_id,
            meta={**metadata, **changes},
            params={
                **model.params.model_dump(),
                **ModelParameterPatch.model_validate(params or {}).model_dump(exclude_unset=True),
            },
            expected_version_id=version_id or _model_versions(__metadata__).get(id, model.version_id),
            commit_message=commit_message,
        )
        return _saved_model(await update_model_by_id(__request__, form, user, None), __metadata__)
    except Exception as error:
        return JSONCodec.dumps({'error': getattr(error, 'detail', str(error))})
