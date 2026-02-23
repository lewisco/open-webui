<script lang="ts">
	import {
		listSharePointItems,
		type SharePointBrowserItem,
		type SharePointSelectedItem
	} from '$lib/apis/sharepoint';
	import Modal from '$lib/components/common/Modal.svelte';
	import Folder from '$lib/components/icons/Folder.svelte';
	import Document from '$lib/components/icons/Document.svelte';
	import XMark from '$lib/components/icons/XMark.svelte';

	import { createEventDispatcher, getContext } from 'svelte';
	import { toast } from 'svelte-sonner';

	const i18n = getContext('i18n');
	const dispatch = createEventDispatcher();

	export let show = false;
	export let driveId = '';
	export let initialSelectedItems: SharePointSelectedItem[] = [];

	let browserItems: SharePointBrowserItem[] = [];
	let loadingItems = false;
	let browserStack: { id: string | null; name: string }[] = [{ id: null, name: 'Root' }];
	let selectedItems: SharePointSelectedItem[] = [];

	// When modal opens, deep-copy initial selections and load root items
	$: if (show && driveId) {
		selectedItems = initialSelectedItems.map((item) => ({ ...item }));
		browserStack = [{ id: null, name: 'Root' }];
		browserItems = [];
		loadBrowserItems(null);
	}

	const loadBrowserItems = async (parentId: string | null) => {
		if (!driveId) return;
		loadingItems = true;
		try {
			browserItems =
				(await listSharePointItems(localStorage.token, driveId, parentId ?? undefined)) ?? [];
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to load items'));
		}
		loadingItems = false;
	};

	const navigateToFolder = async (folderId: string, folderName: string) => {
		browserStack = [...browserStack, { id: folderId, name: folderName }];
		await loadBrowserItems(folderId);
	};

	const navigateToBreadcrumb = async (index: number) => {
		browserStack = browserStack.slice(0, index + 1);
		const current = browserStack[browserStack.length - 1];
		await loadBrowserItems(current.id);
	};

	const getCurrentPath = (): string => {
		return (
			'/' +
			browserStack
				.slice(1)
				.map((b) => b.name)
				.join('/')
		);
	};

	const toggleItemSelection = (item: SharePointBrowserItem) => {
		const itemPath =
			getCurrentPath() === '/' ? `/${item.name}` : `${getCurrentPath()}/${item.name}`;

		const idx = selectedItems.findIndex((s) => s.id === item.id);
		if (idx >= 0) {
			selectedItems = selectedItems.filter((s) => s.id !== item.id);
		} else {
			selectedItems = [
				...selectedItems,
				{
					type: item.isFolder ? 'folder' : 'file',
					id: item.id,
					path: itemPath,
					name: item.name
				}
			];
		}
	};

	const isItemSelected = (itemId: string): boolean => {
		return selectedItems.some((s) => s.id === itemId);
	};

	const countDescendantSelections = (folderName: string): number => {
		const folderPath =
			getCurrentPath() === '/' ? `/${folderName}` : `${getCurrentPath()}/${folderName}`;
		return selectedItems.filter((s) => s.path.startsWith(folderPath + '/')).length;
	};

	const handleConfirm = () => {
		dispatch('confirm', selectedItems);
		show = false;
	};
</script>

<Modal size="md" bind:show>
	<div>
		<!-- Header -->
		<div class="flex justify-between dark:text-gray-100 px-5 pt-4 pb-1.5">
			<div class="text-lg font-medium self-center font-primary">
				{$i18n.t('Select files and folders to sync')}
			</div>
			<button
				class="self-center"
				aria-label={$i18n.t('Close modal')}
				on:click={() => {
					show = false;
				}}
			>
				<XMark className="size-5" />
			</button>
		</div>

		<!-- Body -->
		<div class="px-5 pb-4 dark:text-gray-200">
			<!-- Breadcrumb -->
			<nav
				aria-label={$i18n.t('Breadcrumb')}
				class="flex items-center gap-1 text-xs text-gray-500 mb-2 flex-wrap"
			>
				{#each browserStack as crumb, idx}
					{#if idx > 0}
						<span aria-hidden="true">/</span>
					{/if}
					<button
						class="hover:text-gray-700 dark:hover:text-gray-300"
						type="button"
						on:click={() => navigateToBreadcrumb(idx)}
					>
						{crumb.name}
					</button>
				{/each}
			</nav>

			<!-- Items list -->
			<div class="border border-gray-200 dark:border-gray-700 rounded-lg max-h-80 overflow-y-auto">
				{#if loadingItems}
					<div class="p-3 text-xs text-gray-500 text-center">
						{$i18n.t('Loading...')}
					</div>
				{:else if browserItems.length === 0}
					<div class="p-3 text-xs text-gray-500 text-center">
						{$i18n.t('No items found')}
					</div>
				{:else}
					{#each browserItems as item}
						<div
							class="flex items-center px-3 py-2 hover:bg-gray-50 dark:hover:bg-gray-800 border-b border-gray-100 dark:border-gray-800 last:border-b-0"
						>
							<label class="flex items-center gap-2 flex-1 cursor-pointer">
								<input
									type="checkbox"
									checked={isItemSelected(item.id)}
									on:change={() => toggleItemSelection(item)}
									class="rounded"
								/>
								<span class="text-xs flex items-center gap-1">
									{#if item.isFolder}
										<Folder className="size-3.5 shrink-0" />
									{:else}
										<Document className="size-3.5 shrink-0" />
									{/if}
									{item.name}
									{#if item.isFolder && !isItemSelected(item.id)}
										{@const descendantCount = countDescendantSelections(item.name)}
										{#if descendantCount > 0}
											<span class="text-[10px] text-gray-400 ml-1">
												({descendantCount} inside)
											</span>
										{/if}
									{/if}
								</span>
							</label>

							{#if item.isFolder}
								<button
									class="text-xs text-gray-500 hover:text-gray-700 dark:hover:text-gray-300 ml-2"
									type="button"
									on:click={() => navigateToFolder(item.id, item.name)}
								>
									{$i18n.t('Browse')}
								</button>
							{:else}
								<span class="text-xs text-gray-400">
									{item.size > 1048576
										? `${(item.size / 1048576).toFixed(1)} MB`
										: `${(item.size / 1024).toFixed(0)} KB`}
								</span>
							{/if}
						</div>
					{/each}
				{/if}
			</div>
		</div>

		<!-- Footer -->
		<div class="flex items-center justify-between px-5 pb-4">
			<div class="text-xs text-gray-500">
				{#if selectedItems.length > 0}
					{selectedItems.length} {$i18n.t('item(s) selected')}
				{:else}
					{$i18n.t('No items selected')}
				{/if}
			</div>
			<button
				class="px-3.5 py-1.5 text-sm font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full"
				type="button"
				on:click={handleConfirm}
			>
				{$i18n.t('Done')}
			</button>
		</div>
	</div>
</Modal>
