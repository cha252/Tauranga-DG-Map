#!/usr/bin/env node
// Reads the job spreadsheet, geocodes any addresses not already in
// geocode-cache.json, prunes entries for addresses no longer present in
// the spreadsheet, and writes the updated cache back out.
//
// This mirrors the exact matching/geocoding logic used in index.html so the
// cache this produces is 100% compatible with what the page would produce
// itself via the "Export cache.json" button — this script just automates
// that step in CI.
//
// Requires Node 18+ (for built-in fetch) and the "xlsx" package.

import { readFileSync, existsSync, writeFileSync } from 'fs';
import * as XLSX from 'xlsx';

const SPREADSHEET_CANDIDATES = ['DGS LIST.xlsx', 'DGS LIST.xls', 'DGS LIST.csv'];
const CACHE_FILE = 'geocode-cache.json';

// Roughly Coromandel Peninsula (tip to base), South Waikato, and the
// western Bay of Plenty including Tauranga / Mount Maunganui, with some
// extra margin around each — north of Taupō, on the eastern side of the
// North Island. Keep this in sync with NZ_TARGET_BBOX in index.html.
const NZ_TARGET_BBOX = { west: 174.7, south: -39.2, east: 177.3, north: -36.0 };

// A descriptive User-Agent identifying this script is required by
// Nominatim's usage policy for automated/scripted use.
// https://operations.osmfoundation.org/policies/nominatim/
const NOMINATIM_USER_AGENT = 'DGS-Job-Site-Mapper-CacheUpdater/1.0 (github actions)';

function findKey(row, candidates){
  const keys = Object.keys(row);
  for (const c of candidates){
    const hit = keys.find(k => k.trim().toLowerCase() === c);
    if (hit) return hit;
  }
  return null;
}

function inRegion(lat, lng){
  return lat >= NZ_TARGET_BBOX.south && lat <= NZ_TARGET_BBOX.north &&
         lng >= NZ_TARGET_BBOX.west && lng <= NZ_TARGET_BBOX.east;
}

function cacheKey(address){
  return address.trim().toLowerCase();
}

function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }

async function geocode(address){
  const bbox = NZ_TARGET_BBOX;
  const viewbox = `${bbox.west},${bbox.north},${bbox.east},${bbox.south}`;
  const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=nz'
    + `&viewbox=${viewbox}&bounded=1`
    + `&q=${encodeURIComponent(address)}`;

  const res = await fetch(url, {
    headers: {
      'Accept': 'application/json',
      'User-Agent': NOMINATIM_USER_AGENT
    }
  });
  if (!res.ok) throw new Error(`geocode request failed (${res.status})`);
  const data = await res.json();

  if (!data.length) return null;
  const lat = parseFloat(data[0].lat);
  const lng = parseFloat(data[0].lon);
  return inRegion(lat, lng) ? { lat, lng } : 'out-of-region';
}

async function main(){
  const spreadsheetName = SPREADSHEET_CANDIDATES.find(existsSync);
  if (!spreadsheetName){
    console.log(`No spreadsheet found (looked for: ${SPREADSHEET_CANDIDATES.join(', ')}). Nothing to do.`);
    return;
  }
  console.log(`Reading ${spreadsheetName}...`);

  const wb = XLSX.readFile(spreadsheetName);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (!rows.length){
    console.log('Spreadsheet is empty. Nothing to do.');
    return;
  }

  const addrKey = findKey(rows[0], ['site', 'site address', 'street address', 'location', 'address']);
  if (!addrKey){
    console.error('Could not find a "Site" (or "Address") column in the spreadsheet — skipping.');
    process.exitCode = 1;
    return;
  }

  const addresses = [...new Set(
    rows.map(r => String(r[addrKey]).trim()).filter(Boolean)
  )];

  let cache = {};
  if (existsSync(CACHE_FILE)){
    try {
      cache = JSON.parse(readFileSync(CACHE_FILE, 'utf8'));
    } catch(e){
      console.warn(`Could not parse existing ${CACHE_FILE}, starting fresh.`);
    }
  }

  const needsLookup = (address) => {
    const key = cacheKey(address);
    if (!Object.prototype.hasOwnProperty.call(cache, key)) return true;
    // No coordinates were stored for out-of-region misses, so if the region
    // boundary has since changed, these need a fresh lookup to find out.
    // Cached coordinate objects don't need re-fetching — the page itself
    // re-checks those against the current boundary at render time.
    return cache[key] === 'out-of-region';
  };
  const toLookup = addresses.filter(needsLookup);
  console.log(`${addresses.length} unique site addresses, ${toLookup.length} new to geocode.`);

  let changed = false;
  for (let i = 0; i < toLookup.length; i++){
    const address = toLookup[i];
    try {
      const result = await geocode(address);
      cache[cacheKey(address)] = result;
      changed = true;
      console.log(`[${i + 1}/${toLookup.length}] ${address} -> ${result ? (result === 'out-of-region' ? 'out of region' : `${result.lat}, ${result.lng}`) : 'not found'}`);
    } catch(err){
      console.warn(`[${i + 1}/${toLookup.length}] ${address} -> lookup failed (${err.message}), leaving uncached for retry next run`);
    }
    // Respect Nominatim's 1 request/second usage policy.
    if (i < toLookup.length - 1) await sleep(1000);
  }

  // Prune entries for addresses no longer present in the spreadsheet, so
  // the cache doesn't accumulate dead weight as jobs come and go.
  const currentKeys = new Set(addresses.map(cacheKey));
  const beforeCount = Object.keys(cache).length;
  for (const key of Object.keys(cache)){
    if (!currentKeys.has(key)) delete cache[key];
  }
  const prunedCount = beforeCount - Object.keys(cache).length;
  if (prunedCount > 0){
    changed = true;
    console.log(`Pruned ${prunedCount} entr${prunedCount === 1 ? 'y' : 'ies'} no longer referenced by the spreadsheet.`);
  }

  if (!changed){
    console.log('No new entries — cache is already up to date.');
    return;
  }

  const sorted = Object.fromEntries(Object.keys(cache).sort().map(k => [k, cache[k]]));
  writeFileSync(CACHE_FILE, JSON.stringify(sorted, null, 2) + '\n');
  console.log(`Wrote ${Object.keys(sorted).length} entries to ${CACHE_FILE}.`);
}

main().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
