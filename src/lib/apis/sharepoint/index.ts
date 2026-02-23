import { WEBUI_API_BASE_URL } from '$lib/constants';
import { splitStream } from '$lib/utils';

const SHAREPOINT_API_BASE = `${WEBUI_API_BASE_URL}/sharepoint`;

// ---- Types ----

export interface SharePointConfig {
	ENABLE_SHAREPOINT_SYNC: boolean;
	SHAREPOINT_TENANT_ID: string;
	SHAREPOINT_CLIENT_ID: string;
	SHAREPOINT_CLIENT_SECRET: string;
	SHAREPOINT_SYNC_INTERVAL: number;
}

export interface SharePointSite {
	id: string;
	site_id: string;
	drive_id: string;
	site_url: string | null;
	site_name: string | null;
	drive_name: string | null;
	selected_items: SharePointSelectedItem[] | null;
	sync_all: boolean;
	display_name: string | null;
	sync_enabled: boolean;
	kb_id: string | null;
	kb_name: string | null;
	sync_mode: string;
	delta_link: string | null;
	last_sync_at: number | null;
	sync_status: string;
	sync_error: string | null;
	file_count: number;
	error_count: number;
	excluded_count: number;
	created_at: number;
	updated_at: number;
}

export interface SharePointFile {
	id: string;
	site_config_id: string;
	sp_item_id: string;
	owui_file_id: string | null;
	filename: string | null;
	sp_item_path: string | null;
	sp_etag: string | null;
	sp_last_modified: string | null;
	allowed_users: string[] | null;
	allowed_groups: string[] | null;
	sync_status: string;
	sync_error: string | null;
	excluded: boolean;
	created_at: number;
	updated_at: number;
}

export interface SharePointDrive {
	id: string;
	name: string;
	driveType: string;
}

export interface SharePointBrowserItem {
	id: string;
	name: string;
	size: number;
	isFolder: boolean;
	childCount: number;
	mimeType: string;
	lastModifiedDateTime: string;
}

export interface SharePointSelectedItem {
	type: string;
	id: string;
	path: string;
	name: string;
}

export interface SharePointSyncEvent {
	type: string;
	site_id?: string;
	site_name?: string;
	total_items?: number;
	current?: number;
	total?: number;
	filename?: string;
	action?: string;
	stats?: SharePointSyncStats;
	error?: string;
	page?: number;
	items_found?: number;
}

export interface SharePointSyncStats {
	added: number;
	updated: number;
	skipped: number;
	deleted: number;
	errors: number;
}

export interface SharePointSyncStatus {
	sync_status: string;
	progress: { current: number; total: number; filename: string } | null;
}

export interface SharePointResolvedSite {
	site_id: string;
	site_name: string;
	web_url: string;
}

// ---- Config ----

export const getSharePointConfig = async (token: string): Promise<SharePointConfig> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/config`, {
		method: 'GET',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const updateSharePointConfig = async (
	token: string,
	config: Partial<SharePointConfig>
): Promise<SharePointConfig> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/config`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify(config)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Browse ----

export const resolveSharePointSite = async (
	token: string,
	url: string
): Promise<SharePointResolvedSite> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/browse/resolve`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify({ url })
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const listSharePointDrives = async (
	token: string,
	siteId: string
): Promise<SharePointDrive[]> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/browse/${encodeURIComponent(siteId)}/drives`, {
		method: 'GET',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const listSharePointItems = async (
	token: string,
	driveId: string,
	parentId?: string
): Promise<SharePointBrowserItem[]> => {
	let error = null;
	const params = parentId ? `?parent_id=${encodeURIComponent(parentId)}` : '';

	const res = await fetch(
		`${SHAREPOINT_API_BASE}/browse/drives/${encodeURIComponent(driveId)}/items${params}`,
		{
			method: 'GET',
			headers: {
				Accept: 'application/json',
				'Content-Type': 'application/json',
				authorization: `Bearer ${token}`
			}
		}
	)
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Sites ----

export const getSharePointSites = async (token: string): Promise<SharePointSite[]> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites`, {
		method: 'GET',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const addSharePointSite = async (
	token: string,
	siteData: Record<string, unknown>
): Promise<SharePointSite> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify(siteData)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const updateSharePointSite = async (
	token: string,
	siteId: string,
	data: Record<string, unknown>
): Promise<SharePointSite> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}`, {
		method: 'PUT',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify(data)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const deleteSharePointSite = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}`, {
		method: 'DELETE',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Site Files ----

export const getSharePointSiteFiles = async (
	token: string,
	siteId: string
): Promise<SharePointFile[]> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}/files`, {
		method: 'GET',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Retry ----

export const retrySharePointErrors = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/retry/${encodeURIComponent(siteId)}`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Cancel Sync ----

export const cancelSharePointSync = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sync/${encodeURIComponent(siteId)}/cancel`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

// ---- Sync ----

export const triggerSharePointSync = async (
	token: string,
	siteId?: string,
	force: boolean = false,
	clearExclusions: boolean = false
) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sync`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify({
			site_id: siteId || null,
			force,
			clear_exclusions: clearExclusions
		})
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const getSharePointSyncStatus = async (
	token: string,
	siteId: string
): Promise<SharePointSyncStatus | null> => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sync/${encodeURIComponent(siteId)}/status`, {
		method: 'GET',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			error = err.detail;
			console.error(err);
			return null;
		});

	if (error) {
		throw error;
	}
	return res;
};

export const triggerSharePointSyncStream = async (
	token: string,
	siteId?: string,
	force: boolean = false,
	clearExclusions: boolean = false,
	onEvent?: (event: SharePointSyncEvent) => void,
	signal?: AbortSignal
): Promise<SharePointSyncEvent | null> => {
	const res = await fetch(`${SHAREPOINT_API_BASE}/sync/stream`, {
		method: 'POST',
		headers: {
			Accept: 'application/json',
			'Content-Type': 'application/json',
			authorization: `Bearer ${token}`
		},
		body: JSON.stringify({
			site_id: siteId || null,
			force,
			clear_exclusions: clearExclusions
		}),
		signal
	});

	if (!res.ok) {
		const err = await res.json();
		throw err.detail || 'Sync stream failed';
	}

	const reader = res
		.body!.pipeThrough(new TextDecoderStream())
		.pipeThrough(splitStream('\n'))
		.getReader();

	let lastEvent: SharePointSyncEvent | null = null;

	for (;;) {
		const { value, done } = await reader.read();
		if (done) break;

		const lines = value.split('\n');
		for (const line of lines) {
			if (!line.startsWith('data: ')) continue;
			const data = line.slice(6);
			if (data === '[DONE]') return lastEvent;

			try {
				const event = JSON.parse(data);
				lastEvent = event;
				onEvent?.(event);
			} catch {
				// skip malformed lines
			}
		}
	}

	return lastEvent;
};
