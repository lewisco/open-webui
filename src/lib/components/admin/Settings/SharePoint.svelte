<script lang="ts">
	import {
		getSharePointConfig,
		updateSharePointConfig,
		getSharePointSites,
		addSharePointSite,
		deleteSharePointSite,
		updateSharePointSite,
		resolveSharePointSite,
		listSharePointDrives,
		listSharePointItems,
		triggerSharePointSync,
		retrySharePointErrors,
		getSharePointSiteFiles
	} from '$lib/apis/sharepoint';

	import Switch from '$lib/components/common/Switch.svelte';
	import SensitiveInput from '$lib/components/common/SensitiveInput.svelte';
	import Tooltip from '$lib/components/common/Tooltip.svelte';

	import { onMount, getContext } from 'svelte';
	import { toast } from 'svelte-sonner';

	const i18n = getContext('i18n');

	// Config state
	let enableSharePointSync = false;
	let tenantId = '';
	let clientId = '';
	let clientSecret = '';
	let syncInterval = 900;

	// Sites list
	let sites: any[] = [];
	let loadingSites = false;

	// Add site flow
	let showAddSite = false;
	let siteUrl = '';
	let resolving = false;
	let resolvedSite: { site_id: string; site_name: string; web_url: string } | null = null;

	let drives: { id: string; name: string; driveType: string }[] = [];
	let selectedDriveId = '';
	let selectedDriveName = '';
	let loadingDrives = false;

	// File browser (shared between add + edit flows)
	let browserDriveId = '';
	let browserItems: any[] = [];
	let loadingItems = false;
	let browserStack: { id: string | null; name: string }[] = [{ id: null, name: 'Root' }];
	let selectedItems: { type: string; id: string; path: string; name: string }[] = [];

	let kbName = '';
	let newSiteSyncMode = 'none';
	let addingSite = false;

	// Sync state
	let syncingSiteId: string | null = null;

	// Files panel state
	let expandedFilesSiteId: string | null = null;
	let siteFiles: Record<string, any[]> = {};
	let loadingFiles: Record<string, boolean> = {};

	// Edit site state
	let editingSiteId: string | null = null;
	let editSyncMode = 'none';
	let savingEdit = false;

	export async function submit() {
		await updateSharePointConfig(localStorage.token, {
			ENABLE_SHAREPOINT_SYNC: enableSharePointSync,
			SHAREPOINT_TENANT_ID: tenantId,
			SHAREPOINT_CLIENT_ID: clientId,
			SHAREPOINT_CLIENT_SECRET: clientSecret,
			SHAREPOINT_SYNC_INTERVAL: syncInterval
		});
	}

	const loadSites = async () => {
		loadingSites = true;
		try {
			sites = (await getSharePointSites(localStorage.token)) ?? [];
		} catch (e) {
			console.error(e);
		}
		loadingSites = false;
	};

	const handleResolve = async () => {
		if (!siteUrl.trim()) return;
		resolving = true;
		resolvedSite = null;
		drives = [];
		selectedDriveId = '';
		browserDriveId = '';
		browserItems = [];
		browserStack = [{ id: null, name: 'Root' }];
		selectedItems = [];

		try {
			resolvedSite = await resolveSharePointSite(localStorage.token, siteUrl.trim());
			if (resolvedSite) {
				kbName = `SharePoint: ${resolvedSite.site_name}`;
				await loadDrives();
			}
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to resolve site');
		}
		resolving = false;
	};

	const loadDrives = async () => {
		if (!resolvedSite) return;
		loadingDrives = true;
		try {
			drives = (await listSharePointDrives(localStorage.token, resolvedSite.site_id)) ?? [];
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to load drives');
		}
		loadingDrives = false;
	};

	const handleDriveSelect = async (e: Event) => {
		const target = e.target as HTMLSelectElement;
		selectedDriveId = target.value;
		browserDriveId = selectedDriveId;
		const drive = drives.find((d) => d.id === selectedDriveId);
		selectedDriveName = drive?.name ?? '';
		browserStack = [{ id: null, name: 'Root' }];
		selectedItems = [];
		await loadBrowserItems(null);
	};

	const loadBrowserItems = async (parentId: string | null) => {
		if (!browserDriveId) return;
		loadingItems = true;
		try {
			browserItems =
				(await listSharePointItems(
					localStorage.token,
					browserDriveId,
					parentId ?? undefined
				)) ?? [];
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to load items');
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
		return '/' + browserStack.slice(1).map((b) => b.name).join('/');
	};

	const toggleItemSelection = (item: any) => {
		const itemPath =
			getCurrentPath() === '/'
				? `/${item.name}`
				: `${getCurrentPath()}/${item.name}`;

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

	const handleAddSite = async () => {
		if (!resolvedSite || !selectedDriveId) return;
		addingSite = true;

		try {
			const siteData = {
				site_url: siteUrl.trim(),
				site_id: resolvedSite.site_id,
				site_name: resolvedSite.site_name,
				drive_id: selectedDriveId,
				drive_name: selectedDriveName,
				selected_items: selectedItems.length > 0 ? selectedItems : null,
				kb_name: kbName || `SharePoint: ${resolvedSite.site_name}`,
				sync_mode: newSiteSyncMode
			};

			await addSharePointSite(localStorage.token, siteData);
			toast.success($i18n.t('Site added successfully'));

			// Reset form
			showAddSite = false;
			siteUrl = '';
			resolvedSite = null;
			drives = [];
			selectedDriveId = '';
			browserDriveId = '';
			browserItems = [];
			browserStack = [{ id: null, name: 'Root' }];
			selectedItems = [];
			kbName = '';
			newSiteSyncMode = 'none';

			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to add site');
		}
		addingSite = false;
	};

	const handleDeleteSite = async (siteId: string) => {
		if (!confirm($i18n.t('Are you sure you want to remove this site?'))) return;
		try {
			await deleteSharePointSite(localStorage.token, siteId);
			toast.success($i18n.t('Site removed'));
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to remove site');
		}
	};

	const handleSync = async (siteId: string, force = false, clearExcl = false) => {
		syncingSiteId = siteId;
		try {
			const result = await triggerSharePointSync(localStorage.token, siteId, force, clearExcl);
			if (result) {
				const siteResult = result.sites?.[0];
				if (siteResult?.error) {
					toast.error(`Sync error: ${siteResult.error}`);
				} else if (siteResult?.stats) {
					const s = siteResult.stats;
					toast.success(
						`Synced: ${s.added} added, ${s.updated} updated, ${s.skipped} skipped, ${s.deleted} deleted, ${s.errors} errors`
					);
				}
			}
			// Clear cached files so they reload with fresh data
			delete siteFiles[siteId];
			siteFiles = siteFiles;
			expandedFilesSiteId = null;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Sync failed');
		}
		syncingSiteId = null;
	};

	const handleRetryErrors = async (siteId: string) => {
		syncingSiteId = siteId;
		try {
			const result = await retrySharePointErrors(localStorage.token, siteId);
			if (result?.stats) {
				const s = result.stats;
				toast.success(
					`Retry: ${s.succeeded} succeeded, ${s.failed} failed out of ${s.retried} retried`
				);
			}
			delete siteFiles[siteId];
			siteFiles = siteFiles;
			expandedFilesSiteId = null;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Retry failed');
		}
		syncingSiteId = null;
	};

	const toggleFiles = async (siteId: string) => {
		if (expandedFilesSiteId === siteId) {
			expandedFilesSiteId = null;
			return;
		}
		expandedFilesSiteId = siteId;
		if (!siteFiles[siteId]) {
			loadingFiles[siteId] = true;
			loadingFiles = loadingFiles;
			try {
				const files = await getSharePointSiteFiles(localStorage.token, siteId);
				siteFiles[siteId] = files ?? [];
				siteFiles = siteFiles;
			} catch (e) {
				console.error(e);
				siteFiles[siteId] = [];
				siteFiles = siteFiles;
			}
			loadingFiles[siteId] = false;
			loadingFiles = loadingFiles;
		}
	};

	// ---- Edit site helpers ----

	const startEditSite = async (site: any) => {
		showAddSite = false;
		editingSiteId = site.id;
		editSyncMode = site.sync_mode || 'none';
		browserDriveId = site.drive_id;
		selectedItems = site.selected_items ? [...site.selected_items] : [];
		browserStack = [{ id: null, name: 'Root' }];
		browserItems = [];
		await loadBrowserItems(null);
	};

	const cancelEditSite = () => {
		editingSiteId = null;
		browserDriveId = '';
		browserItems = [];
		browserStack = [{ id: null, name: 'Root' }];
		selectedItems = [];
	};

	const handleSaveEdit = async (siteId: string) => {
		savingEdit = true;
		try {
			await updateSharePointSite(localStorage.token, siteId, {
				selected_items: selectedItems.length > 0 ? selectedItems : null,
				sync_mode: editSyncMode
			});
			toast.success($i18n.t('Site updated'));
			editingSiteId = null;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : 'Failed to update site');
		}
		savingEdit = false;
	};

	const formatTimestamp = (ts: number | null): string => {
		if (!ts) return 'Never';
		return new Date(ts * 1000).toLocaleString();
	};

	onMount(async () => {
		try {
			const res = await getSharePointConfig(localStorage.token);
			if (res) {
				enableSharePointSync = res.ENABLE_SHAREPOINT_SYNC;
				tenantId = res.SHAREPOINT_TENANT_ID;
				clientId = res.SHAREPOINT_CLIENT_ID;
				clientSecret = res.SHAREPOINT_CLIENT_SECRET;
				syncInterval = res.SHAREPOINT_SYNC_INTERVAL;
			}
		} catch (e) {
			console.error(e);
		}

		await loadSites();
	});
</script>

<div class="space-y-3 text-sm">
	<div class="mb-2.5 flex w-full justify-between">
		<div class="self-center text-xs font-medium">
			{$i18n.t('SharePoint Sync')}
		</div>
		<div class="flex items-center relative">
			<Switch bind:state={enableSharePointSync} />
		</div>
	</div>

	{#if enableSharePointSync}
		<div class="mb-2.5 flex flex-col w-full">
			<div class="mb-1 text-xs font-medium">{$i18n.t('Tenant ID')}</div>
			<input
				class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
				type="text"
				placeholder={$i18n.t('Enter Azure AD Tenant ID')}
				bind:value={tenantId}
				autocomplete="off"
			/>
		</div>

		<div class="mb-2.5 flex flex-col w-full">
			<div class="mb-1 text-xs font-medium">{$i18n.t('Client ID')}</div>
			<input
				class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
				type="text"
				placeholder={$i18n.t('Enter App Registration Client ID')}
				bind:value={clientId}
				autocomplete="off"
			/>
		</div>

		<div class="mb-2.5 flex flex-col w-full">
			<div class="mb-1 text-xs font-medium">{$i18n.t('Client Secret')}</div>
			<div class="flex w-full">
				<div class="flex-1">
					<SensitiveInput
						placeholder={$i18n.t('Enter Client Secret')}
						bind:value={clientSecret}
					/>
				</div>
			</div>
		</div>

		<div class="mb-2.5 flex flex-col w-full">
			<div class="mb-1 text-xs font-medium">
				{$i18n.t('Sync Interval (seconds)')}
			</div>
			<input
				class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
				type="number"
				min="60"
				placeholder="900"
				bind:value={syncInterval}
			/>
		</div>

		<!-- Sites Section -->
		<div class="mb-3">
			<div class="mt-0.5 mb-2.5 flex items-center justify-between">
				<div class="text-xs font-medium">{$i18n.t('Sites')}</div>
				<button
					class="px-3 py-1 text-xs font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full"
					type="button"
					on:click={() => {
						if (!showAddSite) {
							editingSiteId = null;
							browserDriveId = '';
							browserItems = [];
							browserStack = [{ id: null, name: 'Root' }];
							selectedItems = [];
						}
						showAddSite = !showAddSite;
					}}
				>
					{showAddSite ? $i18n.t('Cancel') : $i18n.t('Add Site')}
				</button>
			</div>

			<hr class="border-gray-100/30 dark:border-gray-850/30 my-2" />

			<!-- Add Site Flow -->
			{#if showAddSite}
				<div class="mb-4 p-3 rounded-lg border border-gray-200 dark:border-gray-700">
					<!-- Step 1: URL Input -->
					<div class="mb-3">
						<div class="mb-1 text-xs font-medium">{$i18n.t('SharePoint Site URL')}</div>
						<div class="flex gap-2">
							<input
								class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
								type="url"
								placeholder="https://contoso.sharepoint.com/sites/MySite"
								bind:value={siteUrl}
								autocomplete="off"
							/>
							<button
								class="px-3 py-1 text-xs font-medium bg-blue-600 hover:bg-blue-700 text-white rounded-lg disabled:opacity-50"
								type="button"
								disabled={resolving || !siteUrl.trim()}
								on:click={handleResolve}
							>
								{resolving ? $i18n.t('Connecting...') : $i18n.t('Connect')}
							</button>
						</div>
					</div>

					<!-- Step 2: Resolved site + Drive selection -->
					{#if resolvedSite}
						<div class="mb-3 text-xs text-green-600 dark:text-green-400">
							Connected to: <strong>{resolvedSite.site_name}</strong>
						</div>

						<div class="mb-3">
							<div class="mb-1 text-xs font-medium">
								{$i18n.t('Document Library')}
							</div>
							{#if loadingDrives}
								<div class="text-xs text-gray-500">{$i18n.t('Loading drives...')}</div>
							{:else}
								<select
									class="w-full dark:bg-gray-900 rounded-lg px-2 p-1 text-sm bg-transparent outline-hidden"
									on:change={handleDriveSelect}
									bind:value={selectedDriveId}
								>
									<option value="">{$i18n.t('Select a document library')}</option>
									{#each drives as drive}
										<option value={drive.id}>{drive.name} ({drive.driveType})</option>
									{/each}
								</select>
							{/if}
						</div>
					{/if}

					<!-- Step 3: File browser -->
					{#if selectedDriveId}
						<div class="mb-3">
							<div class="mb-1 text-xs font-medium">
								{$i18n.t('Select files and folders to sync')}
							</div>

							<!-- Breadcrumb -->
							<div class="flex items-center gap-1 text-xs text-gray-500 mb-2 flex-wrap">
								{#each browserStack as crumb, idx}
									{#if idx > 0}
										<span>/</span>
									{/if}
									<button
										class="hover:text-blue-600 hover:underline"
										type="button"
										on:click={() => navigateToBreadcrumb(idx)}
									>
										{crumb.name}
									</button>
								{/each}
							</div>

							<!-- Items list -->
							<div
								class="border border-gray-200 dark:border-gray-700 rounded-lg max-h-64 overflow-y-auto"
							>
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
												<span class="text-xs">
													{#if item.isFolder}
														<span class="mr-1">&#128193;</span>
													{:else}
														<span class="mr-1">&#128196;</span>
													{/if}
													{item.name}
												</span>
											</label>

											{#if item.isFolder}
												<button
													class="text-xs text-blue-600 hover:underline ml-2"
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

							{#if selectedItems.length > 0}
								<div class="mt-2 text-xs text-gray-500">
									{selectedItems.length}
									{$i18n.t('item(s) selected')}
								</div>
							{:else}
								<div class="mt-2 text-xs text-gray-400">
									{$i18n.t('No selection = sync everything in the library')}
								</div>
							{/if}
						</div>

						<!-- KB Name -->
						<div class="mb-3">
							<div class="mb-1 text-xs font-medium">
								{$i18n.t('Knowledge Base Name')}
							</div>
							<input
								class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
								type="text"
								placeholder={$i18n.t('Knowledge Base name')}
								bind:value={kbName}
								autocomplete="off"
							/>
						</div>

						<!-- Permission Mode -->
						<div class="mb-3 flex w-full justify-between">
							<div class="self-center text-xs font-medium">
								<Tooltip
									content={$i18n.t(
										'Filter mode restricts KB results based on SharePoint permissions. Requires Entra ID SSO.'
									)}
								>
									{$i18n.t('Permission Mode')}
								</Tooltip>
							</div>
							<div class="flex items-center relative">
								<select
									class="dark:bg-gray-900 w-fit pr-8 rounded-sm px-2 p-1 text-xs bg-transparent outline-hidden text-right"
									bind:value={newSiteSyncMode}
								>
									<option value="none">{$i18n.t('None (open access)')}</option>
									<option value="filter">{$i18n.t('Filter by SharePoint ACLs')}</option>
								</select>
							</div>
						</div>

						<!-- Add button -->
						<div class="flex justify-end">
							<button
								class="px-3 py-1.5 text-xs font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
								type="button"
								disabled={addingSite || !resolvedSite || !selectedDriveId}
								on:click={handleAddSite}
							>
								{addingSite ? $i18n.t('Adding...') : $i18n.t('Add Site')}
							</button>
						</div>
					{/if}
				</div>
			{/if}

			<!-- Sites list -->
			{#if loadingSites}
				<div class="text-xs text-gray-500 text-center py-3">
					{$i18n.t('Loading sites...')}
				</div>
			{:else if sites.length === 0}
				<div class="text-xs text-gray-500 text-center py-3">
					{$i18n.t('No SharePoint sites configured')}
				</div>
			{:else}
				<div class="space-y-2">
					{#each sites as site}
						<div
							class="p-3 rounded-lg border border-gray-200 dark:border-gray-700"
						>
							<div class="flex items-center justify-between mb-2">
								<div>
									<div class="text-sm font-medium">
										{site.site_name || site.site_id}
									</div>
									<div class="text-xs text-gray-500">
										{site.drive_name || site.drive_id} &middot;
										{site.file_count + site.error_count} {$i18n.t('files')}
										{#if site.error_count > 0}
											({site.error_count} {$i18n.t('failed')})
										{/if}
										&middot;
										{$i18n.t('Last sync')}: {formatTimestamp(site.last_sync_at)}
										{#if site.sync_mode === 'filter'}
											&middot; <span class="text-blue-500">{$i18n.t('ACL Filtered')}</span>
										{/if}
									</div>
								</div>
								<div class="flex items-center gap-1">
									{#if site.sync_status === 'syncing' || syncingSiteId === site.id}
										<span class="text-xs text-blue-600">
											{$i18n.t('Syncing...')}
										</span>
									{:else if site.sync_status === 'error'}
										<Tooltip content={site.sync_error || 'Unknown error'}>
											<span class="text-xs text-red-600">
												{$i18n.t('Error')}
											</span>
										</Tooltip>
									{/if}
								</div>
							</div>

							<!-- File status breakdown -->
							<div class="mb-2">
								<button
									class="flex items-center gap-1.5 text-xs text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-300 transition"
									type="button"
									on:click={() => toggleFiles(site.id)}
								>
									<svg
										xmlns="http://www.w3.org/2000/svg"
										viewBox="0 0 16 16"
										fill="currentColor"
										class="size-3 transition-transform {expandedFilesSiteId === site.id ? 'rotate-90' : ''}"
									>
										<path
											fill-rule="evenodd"
											d="M6.22 4.22a.75.75 0 0 1 1.06 0l3.25 3.25a.75.75 0 0 1 0 1.06l-3.25 3.25a.75.75 0 0 1-1.06-1.06L8.94 8 6.22 5.28a.75.75 0 0 1 0-1.06Z"
											clip-rule="evenodd"
										/>
									</svg>
									<span>{site.file_count} {$i18n.t('synced')}</span>
									{#if site.error_count > 0}
										<span class="text-red-500">&middot; {site.error_count} {$i18n.t('error(s)')}</span>
									{/if}
									{#if site.excluded_count > 0}
										<span class="text-gray-400">&middot; {site.excluded_count} {$i18n.t('excluded')}</span>
									{/if}
								</button>

								{#if expandedFilesSiteId === site.id}
									<div
										class="mt-2 rounded-lg border border-gray-100 dark:border-gray-800 overflow-hidden max-h-64 overflow-y-auto"
									>
										{#if loadingFiles[site.id]}
											<div class="p-3 text-xs text-gray-500 text-center">
												{$i18n.t('Loading...')}
											</div>
										{:else if !siteFiles[site.id] || siteFiles[site.id].length === 0}
											<div class="p-3 text-xs text-gray-500 text-center">
												{$i18n.t('No files tracked')}
											</div>
										{:else}
											{#each siteFiles[site.id] as file}
												<div
													class="flex items-start gap-2 px-3 py-2 border-b border-gray-100 dark:border-gray-800 last:border-b-0"
												>
													<!-- Status icon -->
													{#if file.excluded}
														<div class="mt-0.5 shrink-0">
															<svg
																xmlns="http://www.w3.org/2000/svg"
																viewBox="0 0 16 16"
																fill="currentColor"
																class="size-3.5 text-gray-400"
															>
																<path
																	fill-rule="evenodd"
																	d="M3.05 3.05a7 7 0 1 1 9.9 9.9 7 7 0 0 1-9.9-9.9Zm1.627.918a5.5 5.5 0 0 0 7.355 7.355L4.677 3.968ZM11.323 12.032 4.968 5.677a5.5 5.5 0 0 0 6.355 6.355Z"
																	clip-rule="evenodd"
																/>
															</svg>
														</div>
													{:else if file.sync_status === 'error'}
														<div class="mt-0.5 shrink-0">
															<svg
																xmlns="http://www.w3.org/2000/svg"
																viewBox="0 0 16 16"
																fill="currentColor"
																class="size-3.5 text-red-500"
															>
																<path
																	fill-rule="evenodd"
																	d="M6.701 2.25c.577-1 2.02-1 2.598 0l5.196 9a1.5 1.5 0 0 1-1.299 2.25H2.804a1.5 1.5 0 0 1-1.3-2.25l5.197-9ZM8 4a.75.75 0 0 1 .75.75v3a.75.75 0 0 1-1.5 0v-3A.75.75 0 0 1 8 4Zm0 8a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z"
																	clip-rule="evenodd"
																/>
															</svg>
														</div>
													{:else}
														<div class="mt-0.5 shrink-0">
															<svg
																xmlns="http://www.w3.org/2000/svg"
																viewBox="0 0 16 16"
																fill="currentColor"
																class="size-3.5 text-green-500"
															>
																<path
																	fill-rule="evenodd"
																	d="M12.416 3.376a.75.75 0 0 1 .208 1.04l-5 7.5a.75.75 0 0 1-1.154.114l-3-3a.75.75 0 0 1 1.06-1.06l2.353 2.353 4.493-6.74a.75.75 0 0 1 1.04-.207Z"
																	clip-rule="evenodd"
																/>
															</svg>
														</div>
													{/if}

													<!-- File info -->
													<div class="flex-1 min-w-0">
														<div class="flex items-center gap-2">
															<span class="text-xs font-medium truncate">
																{file.filename || 'Unknown'}
															</span>
															{#if file.excluded}
																<span
																	class="shrink-0 px-1.5 py-0.5 text-[10px] rounded bg-gray-100 dark:bg-gray-800 text-gray-500"
																>
																	{$i18n.t('Excluded')}
																</span>
															{:else if file.sync_status === 'error'}
																<span
																	class="shrink-0 px-1.5 py-0.5 text-[10px] rounded bg-red-50 dark:bg-red-900/20 text-red-600 dark:text-red-400"
																>
																	{$i18n.t('Error')}
																</span>
															{:else}
																<span
																	class="shrink-0 px-1.5 py-0.5 text-[10px] rounded bg-green-50 dark:bg-green-900/20 text-green-600 dark:text-green-400"
																>
																	{$i18n.t('Synced')}
																</span>
															{/if}
														</div>
														<div class="text-[11px] text-gray-400 truncate">
															{file.sp_item_path || ''}
															{#if file.sync_error}
																<Tooltip content={file.sync_error}>
																	<span class="text-red-400 dark:text-red-500">
																		&mdash; {file.sync_error.length > 80
																			? file.sync_error.slice(0, 80) + '...'
																			: file.sync_error}
																	</span>
																</Tooltip>
															{:else if file.excluded}
																<span>
																	&mdash; {$i18n.t('Removed from Knowledge Base by user')}
																</span>
															{/if}
														</div>
													</div>
												</div>
											{/each}
										{/if}
									</div>
								{/if}
							</div>

							<div class="flex items-center gap-2 flex-wrap">
								<button
									class="px-2 py-1 text-xs rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-50 dark:hover:bg-gray-800 disabled:opacity-50"
									type="button"
									disabled={syncingSiteId === site.id}
									on:click={() => handleSync(site.id)}
								>
									{$i18n.t('Sync Now')}
								</button>

								{#if site.error_count > 0}
									<button
										class="px-2 py-1 text-xs rounded-lg border border-yellow-300 dark:border-yellow-600 text-yellow-600 dark:text-yellow-400 hover:bg-yellow-50 dark:hover:bg-yellow-900/20 disabled:opacity-50"
										type="button"
										disabled={syncingSiteId === site.id}
										on:click={() => handleRetryErrors(site.id)}
									>
										{$i18n.t('Retry Errors')}
									</button>
								{/if}

								<button
									class="px-2 py-1 text-xs rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-50 dark:hover:bg-gray-800 disabled:opacity-50"
									type="button"
									disabled={syncingSiteId === site.id}
									on:click={() => handleSync(site.id, true, true)}
								>
									{$i18n.t('Force Sync')}
								</button>

								<button
									class="px-2 py-1 text-xs rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-50 dark:hover:bg-gray-800"
									type="button"
									on:click={() => editingSiteId === site.id ? cancelEditSite() : startEditSite(site)}
								>
									{editingSiteId === site.id ? $i18n.t('Cancel Edit') : $i18n.t('Edit')}
								</button>

								<button
									class="px-2 py-1 text-xs rounded-lg border border-red-300 dark:border-red-600 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20"
									type="button"
									on:click={() => handleDeleteSite(site.id)}
								>
									{$i18n.t('Remove')}
								</button>
							</div>

							<!-- Edit panel -->
							{#if editingSiteId === site.id}
								<div class="mt-3 p-3 rounded-lg border border-blue-200 dark:border-blue-800 bg-blue-50/30 dark:bg-blue-900/10">
									<!-- Permission Mode -->
									<div class="mb-3 flex w-full justify-between">
										<div class="self-center text-xs font-medium">
											<Tooltip
												content={$i18n.t(
													'Filter mode restricts KB results based on SharePoint permissions. Requires Entra ID SSO.'
												)}
											>
												{$i18n.t('Permission Mode')}
											</Tooltip>
										</div>
										<div class="flex items-center relative">
											<select
												class="dark:bg-gray-900 w-fit pr-8 rounded-sm px-2 p-1 text-xs bg-transparent outline-hidden text-right"
												bind:value={editSyncMode}
											>
												<option value="none">{$i18n.t('None (open access)')}</option>
												<option value="filter">{$i18n.t('Filter by SharePoint ACLs')}</option>
											</select>
										</div>
									</div>

									<!-- File browser for edit -->
									<div class="mb-3">
										<div class="mb-1 text-xs font-medium">
											{$i18n.t('Select files and folders to sync')}
										</div>

										<!-- Breadcrumb -->
										<div class="flex items-center gap-1 text-xs text-gray-500 mb-2 flex-wrap">
											{#each browserStack as crumb, idx}
												{#if idx > 0}
													<span>/</span>
												{/if}
												<button
													class="hover:text-blue-600 hover:underline"
													type="button"
													on:click={() => navigateToBreadcrumb(idx)}
												>
													{crumb.name}
												</button>
											{/each}
										</div>

										<!-- Items list -->
										<div
											class="border border-gray-200 dark:border-gray-700 rounded-lg max-h-64 overflow-y-auto"
										>
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
															<span class="text-xs">
																{#if item.isFolder}
																	<span class="mr-1">&#128193;</span>
																{:else}
																	<span class="mr-1">&#128196;</span>
																{/if}
																{item.name}
															</span>
														</label>

														{#if item.isFolder}
															<button
																class="text-xs text-blue-600 hover:underline ml-2"
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

										{#if selectedItems.length > 0}
											<div class="mt-2 text-xs text-gray-500">
												{selectedItems.length}
												{$i18n.t('item(s) selected')}
											</div>
										{:else}
											<div class="mt-2 text-xs text-gray-400">
												{$i18n.t('No selection = sync everything in the library')}
											</div>
										{/if}
									</div>

									<!-- Save / Cancel -->
									<div class="flex justify-end gap-2">
										<button
											class="px-3 py-1 text-xs rounded-lg border border-gray-300 dark:border-gray-600 hover:bg-gray-50 dark:hover:bg-gray-800"
											type="button"
											on:click={cancelEditSite}
										>
											{$i18n.t('Cancel')}
										</button>
										<button
											class="px-3 py-1.5 text-xs font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
											type="button"
											disabled={savingEdit}
											on:click={() => handleSaveEdit(site.id)}
										>
											{savingEdit ? $i18n.t('Saving...') : $i18n.t('Save Changes')}
										</button>
									</div>
								</div>
							{/if}
						</div>
					{/each}
				</div>
			{/if}
		</div>
	{/if}
</div>
