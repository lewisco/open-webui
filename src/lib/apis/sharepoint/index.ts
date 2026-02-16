import { WEBUI_API_BASE_URL } from '$lib/constants';

const SHAREPOINT_API_BASE = `${WEBUI_API_BASE_URL}/sharepoint`;

// ---- Config ----

export const getSharePointConfig = async (token: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/config`, {
		method: 'GET',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const updateSharePointConfig = async (token: string, config: Record<string, unknown>) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/config`, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		},
		body: JSON.stringify(config)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

// ---- Browse ----

export const resolveSharePointSite = async (token: string, url: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/browse/resolve`, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		},
		body: JSON.stringify({ url })
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const listSharePointDrives = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/browse/${encodeURIComponent(siteId)}/drives`, {
		method: 'GET',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const listSharePointItems = async (
	token: string,
	driveId: string,
	parentId?: string
) => {
	let error = null;
	const params = parentId ? `?parent_id=${encodeURIComponent(parentId)}` : '';

	const res = await fetch(
		`${SHAREPOINT_API_BASE}/browse/drives/${encodeURIComponent(driveId)}/items${params}`,
		{
			method: 'GET',
			headers: {
				'Content-Type': 'application/json',
				Authorization: `Bearer ${token}`
			}
		}
	)
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

// ---- Sites ----

export const getSharePointSites = async (token: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites`, {
		method: 'GET',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const addSharePointSite = async (token: string, siteData: Record<string, unknown>) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites`, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		},
		body: JSON.stringify(siteData)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const updateSharePointSite = async (
	token: string,
	siteId: string,
	data: Record<string, unknown>
) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}`, {
		method: 'PUT',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		},
		body: JSON.stringify(data)
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

export const deleteSharePointSite = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}`, {
		method: 'DELETE',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

// ---- Site Files ----

export const getSharePointSiteFiles = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/sites/${encodeURIComponent(siteId)}/files`, {
		method: 'GET',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};

// ---- Retry ----

export const retrySharePointErrors = async (token: string, siteId: string) => {
	let error = null;

	const res = await fetch(`${SHAREPOINT_API_BASE}/retry/${encodeURIComponent(siteId)}`, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
		}
	})
		.then(async (res) => {
			if (!res.ok) throw await res.json();
			return res.json();
		})
		.catch((err) => {
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
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
			'Content-Type': 'application/json',
			Authorization: `Bearer ${token}`
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
			console.error(err);
			error = err.detail;
			return null;
		});

	if (error) throw error;
	return res;
};
