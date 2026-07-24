'use strict';

let directoryData = null;
let detailsData = null;
let evidenceData = null;
let packagePaths = null;
let embeddedDetails = null;
let embeddedEvidence = null;

const normalize = value => String(value || '').toLowerCase().trim();
const includes = (value, query) => normalize(value).includes(query);
const formatLocation = row => [row.city, row.state, row.zip].filter(Boolean).join(', ');

async function decodeGzip(buffer) {
  if (typeof DecompressionStream !== 'function') {
    throw new Error('This browser cannot decompress the Atlas data package. Use a current Chrome, Edge, Firefox, or Safari release.');
  }
  const stream = new Blob([buffer]).stream().pipeThrough(new DecompressionStream('gzip'));
  return JSON.parse(await new Response(stream).text());
}

async function fetchPackage(relativePath) {
  const response = await fetch(new URL(relativePath, self.location.href));
  if (!response.ok) throw new Error(`Atlas package failed to load: HTTP ${response.status}`);
  return decodeGzip(await response.arrayBuffer());
}

function compactUnion(row) {
  return {
    id: row.id,
    name: row.name,
    affiliation: row.affiliation,
    designation: row.designation,
    localNumber: row.localNumber,
    level: row.level,
    city: row.city,
    state: row.state,
    zip: row.zip,
    latitude: row.latitude,
    longitude: row.longitude,
    geocodePrecision: row.geocodePrecision,
    website: row.website,
    membership: row.membership,
    finances: row.finances,
    sector: row.sector,
    parent: row.parent,
    background: row.background,
    sources: row.sources,
    publicRecordCoverage: row.publicRecordCoverage,
    updatedAt: row.updatedAt
  };
}

function paginate(rows, page, pageSize) {
  const safeSize = Math.max(1, Math.min(Number(pageSize) || 50, 100));
  const pages = Math.max(1, Math.ceil(rows.length / safeSize));
  const safePage = Math.max(0, Math.min(Number(page) || 0, pages - 1));
  return {
    page: safePage,
    pageSize: safeSize,
    pages,
    total: rows.length,
    rows: rows.slice(safePage * safeSize, (safePage + 1) * safeSize)
  };
}

function searchDirectory(payload) {
  const query = normalize(payload.query);
  const filtered = directoryData.directory.filter(row => {
    if (payload.state && row.state !== payload.state) return false;
    if (payload.affiliation && row.affiliation !== payload.affiliation) return false;
    if (payload.level && row.level !== payload.level) return false;
    if (payload.coverage === 'website' && !row.website) return false;
    if (payload.coverage === 'membership' && !Number.isFinite(row.membership?.latest?.count)) return false;
    if (payload.coverage === 'address' && row.geocodePrecision !== 'address') return false;
    if (!query) return true;
    return [
      row.name,
      row.affiliation,
      row.designation,
      row.localNumber,
      row.city,
      row.state,
      row.zip,
      row.parent?.name,
      ...(row.sector?.tags || [])
    ].some(value => includes(value, query));
  }).sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
  const result = paginate(filtered, payload.page, payload.pageSize);
  return { ...result, rows: result.rows.map(compactUnion) };
}

const toRadians = value => value * Math.PI / 180;
function distanceMiles(aLat, aLon, bLat, bLon) {
  const dLat = toRadians(bLat - aLat);
  const dLon = toRadians(bLon - aLon);
  const lat1 = toRadians(aLat);
  const lat2 = toRadians(bLat);
  const h = Math.sin(dLat / 2) ** 2
    + Math.sin(dLon / 2) ** 2 * Math.cos(lat1) * Math.cos(lat2);
  return 3958.8 * 2 * Math.atan2(Math.sqrt(h), Math.sqrt(1 - h));
}
function bearingDegrees(aLat, aLon, bLat, bLon) {
  const lat1 = toRadians(aLat);
  const lat2 = toRadians(bLat);
  const dLon = toRadians(bLon - aLon);
  const y = Math.sin(dLon) * Math.cos(lat2);
  const x = Math.cos(lat1) * Math.sin(lat2)
    - Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLon);
  return (Math.atan2(y, x) * 180 / Math.PI + 360) % 360;
}

function nearest(payload) {
  const zip = String(payload.zip || '').match(/\d{5}/)?.[0];
  const origin = zip ? directoryData.zctas[zip] : null;
  if (!origin) {
    return { error: 'Enter a valid five-digit ZIP code covered by the Census Gazetteer.' };
  }
  const radius = Math.max(5, Math.min(Number(payload.radius) || 100, 1000));
  const rows = directoryData.directory
    .filter(row => Number.isFinite(row.latitude) && Number.isFinite(row.longitude))
    .map(row => ({
      row,
      distance: distanceMiles(origin[0], origin[1], row.latitude, row.longitude),
      bearing: bearingDegrees(origin[0], origin[1], row.latitude, row.longitude)
    }))
    .filter(item => item.distance <= radius)
    .sort((a, b) => a.distance - b.distance || a.row.name.localeCompare(b.row.name));
  const result = paginate(rows, payload.page, payload.pageSize);
  return {
    ...result,
    zip,
    radius,
    origin,
    compass: rows.slice(0, 24).map(item => ({
      id: item.row.id,
      name: item.row.name,
      distance: item.distance,
      bearing: item.bearing
    })),
    rows: result.rows.map(item => ({
      ...compactUnion(item.row),
      distance: item.distance,
      bearing: item.bearing
    }))
  };
}

async function ensureEvidence() {
  if (evidenceData) return;
  evidenceData = embeddedEvidence
    ? await decodeGzip(embeddedEvidence)
    : await fetchPackage(packagePaths.evidence);
  embeddedEvidence = null;
}

async function ensureDetails() {
  if (detailsData) return;
  detailsData = embeddedDetails
    ? await decodeGzip(embeddedDetails)
    : await fetchPackage(packagePaths.details);
  embeddedDetails = null;
}

function searchOrganizing(payload) {
  const query = normalize(payload.query);
  const rows = evidenceData.organizing.attempts.filter(row => {
    if (payload.state && row.location?.state_or_territory !== payload.state) return false;
    if (payload.outcome && row.outcome?.category !== payload.outcome) return false;
    if (payload.year && !String(row.filed_at || '').startsWith(String(payload.year))) return false;
    if (!query) return true;
    return [
      row.case_name,
      row.case_number,
      row.employer,
      ...(row.petitioner_organizations || []),
      ...(row.involved_labor_organizations || [])
    ].some(value => includes(value, query));
  }).sort((a, b) => String(b.filed_at).localeCompare(String(a.filed_at)) || a.id.localeCompare(b.id));
  const result = paginate(rows, payload.page, payload.pageSize);
  return {
    ...result,
    summary: evidenceData.organizing.summary,
    rows: result.rows.map(row => ({
      id: row.id,
      caseName: row.case_name,
      caseNumber: row.case_number,
      caseUrl: row.case_url,
      employer: row.employer,
      filedAt: row.filed_at,
      closedAt: row.closed_at,
      state: row.location?.state_or_territory || null,
      city: row.location?.city || null,
      status: row.status,
      outcome: row.outcome?.category || null,
      final: row.outcome?.final || false,
      petitionType: row.petition?.type || null,
      newRepresentationAttempt: row.petition?.new_representation_attempt || false,
      petitionerOrganizations: row.petitioner_organizations || [],
      employeesOnPetition: row.unit?.employees_on_petition ?? null,
      voters: row.unit?.voters ?? null
    }))
  };
}

function searchResearch(payload) {
  const query = normalize(payload.query);
  const rows = evidenceData.research.articles.filter(row => {
    if (payload.theme && !(row.themes || []).includes(payload.theme)) return false;
    if (!query) return true;
    return [
      row.title,
      row.journal,
      row.method,
      ...(row.authors || []),
      ...(row.themes || []),
      ...(row.key_findings || [])
    ].some(value => includes(value, query));
  }).sort((a, b) => Number(b.year || 0) - Number(a.year || 0) || a.title.localeCompare(b.title));
  const result = paginate(rows, payload.page, payload.pageSize);
  return {
    ...result,
    summary: evidenceData.research.summary,
    statistics: evidenceData.research.statistics,
    rows: result.rows
  };
}

self.onmessage = async event => {
  const { id, type, payload = {} } = event.data || {};
  try {
    let result;
    if (type === 'init-url') {
      packagePaths = payload.packages;
      directoryData = await fetchPackage(packagePaths.directory);
      result = {
        records: directoryData.directory.length,
        geocoded: directoryData.quality.withCoordinates
      };
    } else if (type === 'init-buffer') {
      packagePaths = payload.packages;
      directoryData = await decodeGzip(payload.directory);
      embeddedDetails = payload.details || null;
      embeddedEvidence = payload.evidence || null;
      result = {
        records: directoryData.directory.length,
        geocoded: directoryData.quality.withCoordinates
      };
    } else if (!directoryData) {
      throw new Error('Atlas directory is still loading.');
    } else if (type === 'directory-search') {
      result = searchDirectory(payload);
    } else if (type === 'near') {
      result = nearest(payload);
    } else if (type === 'load-evidence') {
      await ensureEvidence();
      result = {
        organizing: evidenceData.organizing.attempts.length,
        research: evidenceData.research.articles.length
      };
    } else if (type === 'record-detail') {
      await ensureDetails();
      const row = directoryData.directory.find(item => item.id === payload.id);
      if (!row) throw new Error(`Union record not found: ${payload.id}`);
      result = { ...compactUnion(row), ...(detailsData.records[payload.id] || {}) };
    } else if (type === 'organizing-search') {
      await ensureEvidence();
      result = searchOrganizing(payload);
    } else if (type === 'research-search') {
      await ensureEvidence();
      result = searchResearch(payload);
    } else {
      throw new Error(`Unknown Atlas worker request: ${type}`);
    }
    self.postMessage({ id, ok: true, result });
  } catch (error) {
    self.postMessage({ id, ok: false, error: error?.message || String(error) });
  }
};
