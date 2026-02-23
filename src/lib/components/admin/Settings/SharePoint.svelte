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
		triggerSharePointSync,
		triggerSharePointSyncStream,
		retrySharePointErrors,
		getSharePointSiteFiles,
		cancelSharePointSync,
		getSharePointSyncStatus,
		type SharePointSite,
		type SharePointFile,
		type SharePointDrive,
		type SharePointResolvedSite,
		type SharePointSelectedItem,
		type SharePointSyncEvent
	} from '$lib/apis/sharepoint';

	import { DropdownMenu } from 'bits-ui';
	import { flyAndScale } from '$lib/utils/transitions';
	import Dropdown from '$lib/components/common/Dropdown.svelte';
	import Switch from '$lib/components/common/Switch.svelte';
	import SensitiveInput from '$lib/components/common/SensitiveInput.svelte';
	import Tooltip from '$lib/components/common/Tooltip.svelte';
	import EllipsisHorizontal from '$lib/components/icons/EllipsisHorizontal.svelte';
	import ArrowPath from '$lib/components/icons/ArrowPath.svelte';
	import GarbageBin from '$lib/components/icons/GarbageBin.svelte';
	import Pencil from '$lib/components/icons/Pencil.svelte';
	import Modal from '$lib/components/common/Modal.svelte';
	import XMark from '$lib/components/icons/XMark.svelte';
	import Plus from '$lib/components/icons/Plus.svelte';
	import ChevronRight from '$lib/components/icons/ChevronRight.svelte';
	import DeleteSiteConfirmDialog from '$lib/components/common/ConfirmDialog.svelte';
	import FileBrowserModal from './SharePoint/FileBrowserModal.svelte';

	import { onMount, onDestroy, getContext } from 'svelte';
	import { toast } from 'svelte-sonner';

	const i18n = getContext('i18n');

	// Config state
	let enableSharePointSync = false;
	let tenantId = '';
	let clientId = '';
	let clientSecret = '';
	let syncInterval = 900;

	// Sites list
	let sites: SharePointSite[] = [];
	let loadingSites = false;

	// Add site form
	let showAddForm = false;
	let siteUrl = '';
	let resolving = false;
	let resolvedSite: SharePointResolvedSite | null = null;

	let drives: SharePointDrive[] = [];
	let selectedDriveId = '';
	let selectedDriveName = '';
	let loadingDrives = false;

	// File browser (shared between add + edit flows)
	let browserDriveId = '';
	let selectedItems: SharePointSelectedItem[] = [];

	let kbName = '';
	let newSiteSyncMode = 'none';
	let addingSite = false;
	let syncAll = false;
	let addDisplayName = '';
	let showAddFileBrowser = false;
	let showEditFileBrowser = false;

	// Sync state
	let syncingSiteId: string | null = null;
	let syncProgress: { current: number; total: number; filename: string } | null = null;
	let syncAbortController: AbortController | null = null;

	// Cancel sync state
	let cancellingSync = false;

	// Lifecycle flag to guard async callbacks after unmount
	let destroyed = false;

	// Progress polling for syncs detected after page reload
	let pollInterval: ReturnType<typeof setInterval> | null = null;
	let pollProgress: Record<string, { current: number; total: number; filename: string } | null> =
		{};

	// Files panel state
	let expandedFilesSiteId: string | null = null;
	let siteFiles: Record<string, SharePointFile[]> = {};
	let loadingFiles: Record<string, boolean> = {};

	// Site dropdown menu state
	let showSiteMenu: Record<string, boolean> = {};

	// Edit site state
	let editingSiteId: string | null = null;
	let editSyncMode = 'none';
	let editSyncAll = false;
	let editDisplayName = '';
	let editKbName = '';
	let savingEdit = false;
	let showEditModal = false;

	// Delete confirmation state
	let showDeleteConfirm = false;
	let pendingDeleteSiteId: string | null = null;

	// Clean up add form state on close
	$: if (!showAddForm) {
		siteUrl = '';
		resolvedSite = null;
		drives = [];
		selectedDriveId = '';
		selectedDriveName = '';
		browserDriveId = '';
		selectedItems = [];
		kbName = '';
		newSiteSyncMode = 'none';
		syncAll = false;
		addDisplayName = '';
		showAddFileBrowser = false;
	}

	// Clean up edit modal state on close
	$: if (!showEditModal && editingSiteId !== null) {
		editingSiteId = null;
		browserDriveId = '';
		selectedItems = [];
		showEditFileBrowser = false;
	}

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
			const result = await getSharePointSites(localStorage.token);
			if (destroyed) return;
			sites = result ?? [];
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
		selectedItems = [];

		try {
			resolvedSite = await resolveSharePointSite(localStorage.token, siteUrl.trim());
			if (resolvedSite) {
				kbName = `SharePoint: ${resolvedSite.site_name}`;
				addDisplayName = resolvedSite.site_name;
				await loadDrives();
			}
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to resolve site'));
		}
		resolving = false;
	};

	const loadDrives = async () => {
		if (!resolvedSite) return;
		loadingDrives = true;
		try {
			drives = (await listSharePointDrives(localStorage.token, resolvedSite.site_id)) ?? [];
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to load drives'));
		}
		loadingDrives = false;
	};

	const handleDriveSelect = (e: Event) => {
		const target = e.target as HTMLSelectElement;
		selectedDriveId = target.value;
		const drive = drives.find((d) => d.id === selectedDriveId);
		selectedDriveName = drive?.name ?? '';
		selectedItems = [];
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
				selected_items: syncAll ? null : selectedItems.length > 0 ? selectedItems : null,
				sync_all: syncAll,
				display_name: addDisplayName || null,
				kb_name: kbName || `SharePoint: ${resolvedSite.site_name}`,
				sync_mode: newSiteSyncMode
			};

			await addSharePointSite(localStorage.token, siteData);
			toast.success($i18n.t('Site added successfully'));

			showAddForm = false;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to add site'));
		}
		addingSite = false;
	};

	const handleDeleteSite = (siteId: string) => {
		pendingDeleteSiteId = siteId;
		showDeleteConfirm = true;
	};

	const confirmDeleteSite = async () => {
		if (!pendingDeleteSiteId) return;
		try {
			await deleteSharePointSite(localStorage.token, pendingDeleteSiteId);
			toast.success($i18n.t('Site removed'));
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to remove site'));
		}
		pendingDeleteSiteId = null;
	};

	const handleToggleSyncEnabled = async (site: SharePointSite) => {
		const newValue = !site.sync_enabled;
		try {
			await updateSharePointSite(localStorage.token, site.id, {
				sync_enabled: newValue
			});
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to update site'));
		}
	};

	const handleSync = async (siteId: string, force = false, clearExcl = false) => {
		syncingSiteId = siteId;
		syncProgress = null;
		const controller = new AbortController();
		syncAbortController = controller;
		try {
			const lastEvent = await triggerSharePointSyncStream(
				localStorage.token,
				siteId,
				force,
				clearExcl,
				(event: SharePointSyncEvent) => {
					if (event.type === 'started') {
						syncProgress = { current: 0, total: 0, filename: '' };
					} else if (event.type === 'discovering') {
						syncProgress = {
							current: 0,
							total: 0,
							filename: ''
						};
					} else if (event.type === 'discovery') {
						syncProgress = {
							current: 0,
							total: event.total_items ?? 0,
							filename: ''
						};
					} else if (event.type === 'progress') {
						syncProgress = {
							current: event.current ?? 0,
							total: event.total ?? 0,
							filename: event.filename ?? ''
						};
					} else if (event.type === 'error') {
						toast.error($i18n.t('Sync error: {{error}}', { error: event.error }));
					}
				},
				controller.signal
			);
			if (lastEvent?.type === 'complete' && lastEvent.stats) {
				const s = lastEvent.stats;
				toast.success(
					$i18n.t(
						'Synced: {{added}} added, {{updated}} updated, {{skipped}} skipped, {{deleted}} deleted, {{errors}} errors',
						{
							added: s.added,
							updated: s.updated,
							skipped: s.skipped,
							deleted: s.deleted,
							errors: s.errors
						}
					)
				);
			}
			// Clear cached files so they reload with fresh data
			delete siteFiles[siteId];
			siteFiles = siteFiles;
			expandedFilesSiteId = null;
			await loadSites();
		} catch (e: any) {
			if (e instanceof DOMException && e.name === 'AbortError') {
				// User cancelled — handled silently
			} else {
				toast.error(typeof e === 'string' ? e : $i18n.t('Sync failed'));
			}
		} finally {
			syncingSiteId = null;
			syncProgress = null;
			syncAbortController = null;
		}
	};

	const handleRetryErrors = async (siteId: string) => {
		syncingSiteId = siteId;
		try {
			const result = await retrySharePointErrors(localStorage.token, siteId);
			if (result?.stats) {
				const s = result.stats;
				toast.success(
					$i18n.t('Retry: {{succeeded}} succeeded, {{failed}} failed out of {{retried}} retried', {
						succeeded: s.succeeded,
						failed: s.failed,
						retried: s.retried
					})
				);
			}
			delete siteFiles[siteId];
			siteFiles = siteFiles;
			expandedFilesSiteId = null;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Retry failed'));
		}
		syncingSiteId = null;
	};

	const handleCancelSync = async (siteId: string) => {
		cancellingSync = true;
		// Abort the SSE stream if this is a frontend-initiated sync
		syncAbortController?.abort();
		syncAbortController = null;
		try {
			await cancelSharePointSync(localStorage.token, siteId);
			toast.success($i18n.t('Sync cancelled'));
			syncingSiteId = null;
			syncProgress = null;
			// Clean up polling state for this site
			delete pollProgress[siteId];
			pollProgress = pollProgress;
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to cancel sync'));
		}
		cancellingSync = false;
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

	const startEditSite = (site: SharePointSite) => {
		editingSiteId = site.id;
		editSyncMode = site.sync_mode || 'none';
		editSyncAll = site.sync_all ?? false;
		editDisplayName = site.display_name || site.site_name || '';
		editKbName = site.kb_name || '';
		browserDriveId = site.drive_id;
		selectedItems = site.selected_items ? [...site.selected_items] : [];
		showEditModal = true;
	};

	const cancelEditSite = () => {
		showEditModal = false;
		editingSiteId = null;
		browserDriveId = '';
		selectedItems = [];
		showEditFileBrowser = false;
	};

	const handleSaveEdit = async (siteId: string) => {
		savingEdit = true;
		try {
			await updateSharePointSite(localStorage.token, siteId, {
				selected_items: editSyncAll ? null : selectedItems.length > 0 ? selectedItems : null,
				sync_all: editSyncAll,
				display_name: editDisplayName || null,
				sync_mode: editSyncMode,
				kb_name: editKbName || null
			});
			toast.success($i18n.t('Site updated'));
			cancelEditSite();
			await loadSites();
		} catch (e: any) {
			toast.error(typeof e === 'string' ? e : $i18n.t('Failed to update site'));
		}
		savingEdit = false;
	};

	const startPolling = () => {
		stopPolling();
		pollInterval = setInterval(async () => {
			if (destroyed) return;
			const syncingSites = sites.filter(
				(s) => s.sync_status === 'syncing' && syncingSiteId !== s.id
			);
			if (syncingSites.length === 0) {
				stopPolling();
				return;
			}
			for (const site of syncingSites) {
				try {
					const result = await getSharePointSyncStatus(localStorage.token, site.id);
					if (destroyed) return;
					if (result) {
						if (result.sync_status !== 'syncing') {
							// Sync finished — clear polling for this site and refresh
							delete pollProgress[site.id];
							pollProgress = pollProgress;
							await loadSites();
						} else {
							pollProgress[site.id] = result.progress;
							pollProgress = pollProgress;
						}
					}
				} catch {
					// Ignore polling errors
				}
			}
		}, 3000);
	};

	const stopPolling = () => {
		if (pollInterval) {
			clearInterval(pollInterval);
			pollInterval = null;
		}
	};

	// Start polling whenever sites load and some are syncing (not from this session)
	$: {
		const hasExternalSync = sites.some(
			(s) => s.sync_status === 'syncing' && syncingSiteId !== s.id
		);
		if (hasExternalSync) {
			startPolling();
		} else {
			stopPolling();
			pollProgress = {};
		}
	}

	const formatTimestamp = (ts: number | null): string => {
		if (!ts) return $i18n.t('Never');
		return new Date(ts * 1000).toLocaleString();
	};

	const getSiteDisplayName = (site: SharePointSite): string => {
		return site.display_name || site.site_name || site.site_id;
	};

	onDestroy(() => {
		destroyed = true;
		stopPolling();
	});

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

<DeleteSiteConfirmDialog
	bind:show={showDeleteConfirm}
	title={$i18n.t('Remove Site')}
	message={$i18n.t(
		'This will remove the site configuration, its Knowledge Base, all synced files, and vector embeddings. This action cannot be undone.'
	)}
	on:confirm={confirmDeleteSite}
/>

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
					<SensitiveInput placeholder={$i18n.t('Enter Client Secret')} bind:value={clientSecret} />
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
				<Tooltip content={$i18n.t('Add Site')}>
					<button
						class="px-1"
						type="button"
						on:click={() => {
							showAddForm = !showAddForm;
						}}
					>
						<Plus />
					</button>
				</Tooltip>
			</div>

			<hr class="border-gray-100/30 dark:border-gray-850/30 my-2" />

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
							class="p-3 rounded-lg border border-gray-200 dark:border-gray-700 {site.sync_enabled ===
							false
								? 'opacity-60'
								: ''}"
						>
							<div class="flex items-center justify-between mb-2">
								<div>
									<div class="text-sm font-medium">
										{getSiteDisplayName(site)}
									</div>
									<div class="text-xs text-gray-500">
										{site.drive_name || site.drive_id} &middot;
										{site.file_count + site.error_count}
										{$i18n.t('files')}
										{#if site.error_count > 0}
											({site.error_count} {$i18n.t('failed')})
										{/if}
										&middot;
										{$i18n.t('Last sync')}: {formatTimestamp(site.last_sync_at)}
										{#if site.sync_mode === 'filter'}
											&middot; <span class="text-gray-500">{$i18n.t('ACL Filtered')}</span>
										{/if}
									</div>
								</div>
								<div class="flex items-center gap-2">
									<div class="flex items-center gap-1" aria-live="polite">
										{#if site.sync_enabled === false}
											<span class="text-xs text-gray-500">{$i18n.t('Paused')}</span>
										{:else if site.sync_status === 'syncing' || syncingSiteId === site.id}
											<span
												class="text-xs text-gray-600 dark:text-gray-400 text-right max-w-[200px]"
											>
												{#if syncingSiteId === site.id && syncProgress && syncProgress.total > 0}
													{syncProgress.current}/{syncProgress.total}
													{#if syncProgress.filename}
														&mdash; <span class="inline-block max-w-[120px] truncate align-bottom"
															>{syncProgress.filename}</span
														>
													{/if}
												{:else if syncingSiteId === site.id && syncProgress}
													{syncProgress.filename || $i18n.t('Discovering files...')}
												{:else if pollProgress[site.id]}
													{@const pp = pollProgress[site.id]}
													{pp?.current}/{pp?.total}
													{#if pp?.filename}
														&mdash; <span class="inline-block max-w-[120px] truncate align-bottom"
															>{pp.filename}</span
														>
													{/if}
												{:else}
													{$i18n.t('Syncing...')}
												{/if}
											</span>
										{:else if site.sync_status === 'error'}
											<Tooltip content={site.sync_error || 'Unknown error'}>
												<span class="text-xs text-red-600">
													{$i18n.t('Error')}
												</span>
											</Tooltip>
										{/if}
									</div>
									<Switch
										state={site.sync_enabled !== false}
										on:change={() => handleToggleSyncEnabled(site)}
									/>
								</div>
							</div>

							<!-- File status breakdown -->
							<div class="mb-2">
								<button
									class="flex items-center gap-1.5 text-xs text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-300 transition"
									type="button"
									on:click={() => toggleFiles(site.id)}
								>
									<ChevronRight
										className="size-3 transition-transform {expandedFilesSiteId === site.id
											? 'rotate-90'
											: ''}"
									/>
									<span>{site.file_count} {$i18n.t('synced')}</span>
									{#if site.error_count > 0}
										<span class="text-red-500"
											>&middot; {site.error_count} {$i18n.t('error(s)')}</span
										>
									{/if}
									{#if site.excluded_count > 0}
										<span class="text-gray-400"
											>&middot; {site.excluded_count} {$i18n.t('excluded')}</span
										>
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

							<div class="flex items-center justify-end gap-1">
								{#if site.sync_status === 'syncing' || syncingSiteId === site.id}
									<button
										class="px-3.5 py-1.5 text-sm font-medium bg-red-600 hover:bg-red-700 text-white transition rounded-full disabled:opacity-50"
										type="button"
										disabled={cancellingSync}
										on:click={() => handleCancelSync(site.id)}
									>
										{cancellingSync ? $i18n.t('Cancelling...') : $i18n.t('Cancel Sync')}
									</button>
								{:else}
									<button
										class="px-3.5 py-1.5 text-sm font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
										type="button"
										disabled={syncingSiteId === site.id}
										on:click={() => handleSync(site.id)}
									>
										{$i18n.t('Sync Now')}
									</button>
								{/if}

								<Dropdown bind:show={showSiteMenu[site.id]} align="end">
									<Tooltip content={$i18n.t('More')}>
										<button
											class="self-center w-fit text-sm p-1.5 dark:text-gray-300 dark:hover:text-white hover:bg-black/5 dark:hover:bg-white/5 rounded-xl"
											type="button"
											on:click={(e) => {
												e.stopPropagation();
												showSiteMenu[site.id] = true;
											}}
										>
											<EllipsisHorizontal className="size-5" />
										</button>
									</Tooltip>

									<div slot="content">
										<DropdownMenu.Content
											class="w-full max-w-[170px] rounded-xl p-1 border border-gray-100 dark:border-gray-800 z-50 bg-white dark:bg-gray-850 dark:text-white shadow-sm"
											side="bottom"
											align="end"
											transition={flyAndScale}
										>
											<DropdownMenu.Item
												class="select-none flex gap-2 items-center px-3 py-1.5 text-sm font-medium cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800 rounded-md {syncingSiteId ===
												site.id
													? 'opacity-50 pointer-events-none'
													: ''}"
												on:click={() => {
													showSiteMenu[site.id] = false;
													handleSync(site.id, true, true);
												}}
											>
												<ArrowPath />
												<div class="flex items-center">{$i18n.t('Force Sync')}</div>
											</DropdownMenu.Item>

											{#if site.error_count > 0}
												<DropdownMenu.Item
													class="select-none flex gap-2 items-center px-3 py-1.5 text-sm font-medium cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800 rounded-md {syncingSiteId ===
													site.id
														? 'opacity-50 pointer-events-none'
														: ''}"
													on:click={() => {
														showSiteMenu[site.id] = false;
														handleRetryErrors(site.id);
													}}
												>
													<ArrowPath />
													<div class="flex items-center">{$i18n.t('Retry Errors')}</div>
												</DropdownMenu.Item>
											{/if}

											<DropdownMenu.Separator
												class="my-0.5 border-t border-gray-100 dark:border-gray-800"
											/>

											<DropdownMenu.Item
												class="select-none flex gap-2 items-center px-3 py-1.5 text-sm font-medium cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800 rounded-md"
												on:click={() => {
													showSiteMenu[site.id] = false;
													startEditSite(site);
												}}
											>
												<Pencil />
												<div class="flex items-center">{$i18n.t('Edit')}</div>
											</DropdownMenu.Item>

											<DropdownMenu.Separator
												class="my-0.5 border-t border-gray-100 dark:border-gray-800"
											/>

											<DropdownMenu.Item
												class="select-none flex gap-2 items-center px-3 py-1.5 text-sm font-medium cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800 rounded-md text-red-500"
												on:click={() => {
													showSiteMenu[site.id] = false;
													handleDeleteSite(site.id);
												}}
											>
												<GarbageBin />
												<div class="flex items-center">{$i18n.t('Remove')}</div>
											</DropdownMenu.Item>
										</DropdownMenu.Content>
									</div>
								</Dropdown>
							</div>
						</div>
					{/each}
				</div>
			{/if}
			<!-- Inline Add Site form -->
			{#if showAddForm}
				<div class="p-3 rounded-lg border border-gray-200 dark:border-gray-700">
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
								class="px-3.5 py-1.5 text-sm font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
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
						<div class="mb-3 text-xs text-gray-600 dark:text-gray-400">
							Connected to: <strong>{resolvedSite.web_url || siteUrl}</strong>
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

					<!-- Step 3: Configuration -->
					{#if selectedDriveId}
						<!-- Name -->
						<div class="mb-3">
							<div class="mb-1 text-xs font-medium">
								{$i18n.t('Name')}
							</div>
							<input
								class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
								type="text"
								placeholder={$i18n.t('Name for this site')}
								bind:value={addDisplayName}
								autocomplete="off"
							/>
						</div>

						<!-- Sync All Toggle -->
						<div class="mb-3 flex w-full justify-between">
							<div class="self-center text-xs font-medium">
								<Tooltip
									content={$i18n.t(
										'When enabled, all files in the library are synced. When disabled, only selected files and folders are synced.'
									)}
								>
									{$i18n.t('Sync all files')}
								</Tooltip>
							</div>
							<div class="flex items-center relative">
								<Switch bind:state={syncAll} />
							</div>
						</div>

						<!-- File selection button (hidden when syncAll is on) -->
						{#if !syncAll}
							<div class="mb-3">
								<div class="mb-1 text-xs font-medium">
									{$i18n.t('Files & Folders')}
								</div>
								<button
									class="w-full text-left px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg hover:bg-gray-50 dark:hover:bg-gray-800 transition"
									type="button"
									on:click={() => {
										showAddFileBrowser = true;
									}}
								>
									{#if selectedItems.length > 0}
										<span class="text-gray-700 dark:text-gray-300">
											{selectedItems.length}
											{$i18n.t('item(s) selected')}
										</span>
									{:else}
										<span class="text-gray-400">
											{$i18n.t('Select files and folders...')}
										</span>
									{/if}
								</button>
								{#if selectedItems.length === 0}
									<div class="mt-1.5 text-xs text-gray-400">
										{$i18n.t('No items selected \u2014 nothing will sync')}
									</div>
								{/if}
							</div>
						{/if}

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

						<!-- Actions -->
						<div class="flex justify-end gap-2">
							<button
								class="px-3 py-1.5 text-sm font-medium text-gray-600 dark:text-gray-400 hover:text-gray-800 dark:hover:text-gray-200 transition"
								type="button"
								on:click={() => {
									showAddForm = false;
								}}
							>
								{$i18n.t('Cancel')}
							</button>
							<button
								class="px-3.5 py-1.5 text-sm font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
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
		</div>
	{/if}
</div>

<!-- Edit Site Modal -->
<Modal size="md" bind:show={showEditModal}>
	<div>
		<!-- Header -->
		<div class="flex justify-between dark:text-gray-100 px-5 pt-4 pb-1.5">
			<div class="text-lg font-medium self-center font-primary">
				{$i18n.t('Edit Site')}
			</div>
			<button class="self-center" aria-label={$i18n.t('Close modal')} on:click={cancelEditSite}>
				<XMark className="size-5" />
			</button>
		</div>

		{#if editingSiteId}
			<!-- Body -->
			<div class="px-5 pb-4 dark:text-gray-200">
				<!-- Name -->
				<div class="mb-3">
					<div class="mb-1 text-xs font-medium">
						{$i18n.t('Name')}
					</div>
					<input
						class="flex-1 w-full rounded-lg text-sm bg-transparent outline-hidden"
						type="text"
						placeholder={$i18n.t('Name for this site')}
						bind:value={editDisplayName}
						autocomplete="off"
					/>
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
						bind:value={editKbName}
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
							bind:value={editSyncMode}
						>
							<option value="none">{$i18n.t('None (open access)')}</option>
							<option value="filter">{$i18n.t('Filter by SharePoint ACLs')}</option>
						</select>
					</div>
				</div>

				<!-- Sync All Toggle -->
				<div class="mb-3 flex w-full justify-between">
					<div class="self-center text-xs font-medium">
						<Tooltip
							content={$i18n.t(
								'When enabled, all files in the library are synced. When disabled, only selected files and folders are synced.'
							)}
						>
							{$i18n.t('Sync all files')}
						</Tooltip>
					</div>
					<div class="flex items-center relative">
						<Switch bind:state={editSyncAll} />
					</div>
				</div>

				<!-- File selection button for edit (hidden when editSyncAll is on) -->
				{#if !editSyncAll}
					<div class="mb-3">
						<div class="mb-1 text-xs font-medium">
							{$i18n.t('Files & Folders')}
						</div>
						<button
							class="w-full text-left px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg hover:bg-gray-50 dark:hover:bg-gray-800 transition"
							type="button"
							on:click={() => {
								showEditFileBrowser = true;
							}}
						>
							{#if selectedItems.length > 0}
								<span class="text-gray-700 dark:text-gray-300">
									{selectedItems.length}
									{$i18n.t('item(s) selected')}
								</span>
							{:else}
								<span class="text-gray-400">
									{$i18n.t('Select files and folders...')}
								</span>
							{/if}
						</button>
						{#if selectedItems.length === 0}
							<div class="mt-1.5 text-xs text-gray-400">
								{$i18n.t("Select files or folders to sync, or enable 'Sync all files'")}
							</div>
						{/if}
					</div>
				{/if}
			</div>

			<!-- Footer -->
			<div class="flex justify-end px-5 pb-4">
				<button
					class="px-3.5 py-1.5 text-sm font-medium bg-black hover:bg-gray-900 text-white dark:bg-white dark:text-black dark:hover:bg-gray-100 transition rounded-full disabled:opacity-50"
					type="button"
					disabled={savingEdit}
					on:click={() => editingSiteId && handleSaveEdit(editingSiteId)}
				>
					{savingEdit ? $i18n.t('Saving...') : $i18n.t('Save')}
				</button>
			</div>
		{/if}
	</div>
</Modal>

<!-- File Browser Modal (for Add Site flow) -->
<FileBrowserModal
	bind:show={showAddFileBrowser}
	driveId={selectedDriveId}
	initialSelectedItems={selectedItems}
	on:confirm={(e) => {
		selectedItems = e.detail;
	}}
/>

<!-- File Browser Modal (for Edit Site flow) -->
<FileBrowserModal
	bind:show={showEditFileBrowser}
	driveId={browserDriveId}
	initialSelectedItems={selectedItems}
	on:confirm={(e) => {
		selectedItems = e.detail;
	}}
/>
