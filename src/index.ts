import {
  applyOfflineCatchupToRun,
  createOfflineCatchupNotification,
  getBiomeKeyFromMap,
  getRareMaterialNameForMap,
  getRuntimeMapFromActiveRun,
} from "./sim";

function json(data, init = {}) {
  const headers = new Headers(init.headers || {});
  headers.set("content-type", "application/json; charset=utf-8");
  headers.set("access-control-allow-origin", "*");
  headers.set("access-control-allow-methods", "GET,POST,OPTIONS");
  headers.set("access-control-allow-headers", "content-type");
  return new Response(JSON.stringify(data), {
    ...init,
    headers,
  });
}

function generateId(prefix = "id") {
  return `${prefix}_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
}

function validateUsername(value) {
  const trimmed = String(value || "").trim();

  if (!trimmed) {
    return "Username is required.";
  }
  if (trimmed.length < 3 || trimmed.length > 16) {
    return "Username must be 3-16 characters.";
  }
  if (!/^[A-Za-z0-9_]+$/.test(trimmed)) {
    return "Use letters, numbers, and underscores only.";
  }

  return "";
}

const PROFILE_SCHEMA_VERSION = 1;
const GLOBAL_CHAT_MAX_LENGTH = 100;
const GLOBAL_CHAT_SEND_COOLDOWN_MS = 2000;
const GLOBAL_CHAT_DUPLICATE_WINDOW_MS = 15000;
const OFFLINE_NOTIFICATION_MIN_MS = 10000;
const MAP_CRAFT_FRAGMENT_COST = 25;
const DEFAULT_BOSS_INTERVAL = 20;

const SAMPLE_MAPS = [
  {
    id: "grasslands",
    name: "Grasslands",
    biome: "Grasslands",
    width: 10,
    height: 6,
    maxLeaks: 20,
    bossConfig: { rotation: ["guardian", "broodmother", "titan"], interval: DEFAULT_BOSS_INTERVAL },
    tiles: [
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
      ["wall","wall","wall","wall","path","path","path","wall","wall","wall"],
      ["wall","wall","build","build","path","build","path","wall","wall","wall"],
      ["wall","spawn","path","path","path","wall","path","build","wall","wall"],
      ["wall","wall","wall","build","wall","build","path","path","castle","wall"],
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
    ],
    path: [
      { x: 1, y: 3 },
      { x: 2, y: 3 },
      { x: 3, y: 3 },
      { x: 4, y: 3 },
      { x: 4, y: 2 },
      { x: 4, y: 1 },
      { x: 5, y: 1 },
      { x: 6, y: 1 },
      { x: 6, y: 2 },
      { x: 6, y: 3 },
      { x: 6, y: 4 },
      { x: 7, y: 4 },
      { x: 8, y: 4 },
    ],
  },
  {
    id: "ember_ridge",
    name: "Ember Ridge",
    biome: "Ember Ridge",
    width: 12,
    height: 6,
    maxLeaks: 20,
    bossConfig: { rotation: ["guardian", "broodmother", "titan"], interval: DEFAULT_BOSS_INTERVAL },
    tiles: [
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
      ["wall","spawn","path","path","path","wall","wall","wall","wall","wall","wall","wall"],
      ["wall","wall","wall","build","path","build","build","wall","wall","wall","wall","wall"],
      ["wall","wall","wall","build","path","path","path","path","build","wall","wall","wall"],
      ["wall","wall","wall","wall","wall","wall","build","path","path","path","castle","wall"],
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
    ],
    path: [
      { x: 1, y: 1 },
      { x: 2, y: 1 },
      { x: 3, y: 1 },
      { x: 4, y: 1 },
      { x: 4, y: 2 },
      { x: 4, y: 3 },
      { x: 5, y: 3 },
      { x: 6, y: 3 },
      { x: 7, y: 3 },
      { x: 7, y: 4 },
      { x: 8, y: 4 },
      { x: 9, y: 4 },
      { x: 10, y: 4 },
    ],
  },
  {
    id: "tidelands",
    name: "Tidelands",
    biome: "Tidelands",
    width: 10,
    height: 6,
    maxLeaks: 20,
    bossConfig: { rotation: ["guardian", "broodmother", "titan"], interval: DEFAULT_BOSS_INTERVAL },
    tiles: [
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
      ["wall","wall","build","build","wall","wall","wall","wall","wall","wall"],
      ["wall","path","path","path","path","build","wall","castle","wall","wall"],
      ["wall","path","wall","wall","path","build","wall","path","wall","wall"],
      ["wall","spawn","wall","build","path","path","path","path","build","wall"],
      ["wall","wall","wall","wall","wall","wall","wall","wall","wall","wall"],
    ],
    path: [
      { x: 1, y: 4 },
      { x: 1, y: 3 },
      { x: 1, y: 2 },
      { x: 2, y: 2 },
      { x: 3, y: 2 },
      { x: 4, y: 2 },
      { x: 4, y: 3 },
      { x: 4, y: 4 },
      { x: 5, y: 4 },
      { x: 6, y: 4 },
      { x: 7, y: 4 },
      { x: 7, y: 3 },
      { x: 7, y: 2 },
    ],
  },
];

function getBaseMapForRuntimeMap(runtimeMap) {
  const candidateIds = [
    runtimeMap?.sourceMapId,
    runtimeMap?.mapId,
    runtimeMap?.id,
  ]
    .filter(Boolean)
    .map((value) => String(value));

  return SAMPLE_MAPS.find((mapItem) => candidateIds.includes(mapItem.id)) || null;
}

function getRuntimeMapTileType(runtimeMap, x, y) {
  const directRow = runtimeMap?.tiles?.[y];
  const directTile = Array.isArray(directRow) ? directRow[x] : null;
  if (directTile) return directTile;

  const baseMap = getBaseMapForRuntimeMap(runtimeMap);
  const baseRow = baseMap?.tiles?.[y];
  return Array.isArray(baseRow) ? baseRow[x] : null;
}


function getUnlockedCraftableBiomes(profile) {
  const unlockedMaps = SAMPLE_MAPS.filter((mapItem) =>
    (profile?.unlocks?.maps || []).includes(mapItem.id)
  );

  const uniqueBiomes = [];
  unlockedMaps.forEach((mapItem) => {
    if (!uniqueBiomes.includes(mapItem.biome)) uniqueBiomes.push(mapItem.biome);
  });

  return uniqueBiomes.length ? uniqueBiomes : ["Grasslands"];
}

function randomBetween(min, max, precision = 2) {
  const value = min + Math.random() * (max - min);
  return Number(value.toFixed(precision));
}

function randomIntBetween(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function rollCraftedMapDangerModifiers() {
  const possible = [
    { key: "enemySpeed", label: "Enemy Speed", value: randomBetween(1.1, 1.5) },
    { key: "enemyHp", label: "Enemy HP", value: randomBetween(1.1, 1.75) },
    { key: "bossInterval", label: "Boss Interval", value: randomIntBetween(12, 18) },
  ];

  const count = Math.random() < 0.5 ? 1 : 2;
  const rolled = [];

  while (rolled.length < count && possible.length) {
    const index = Math.floor(Math.random() * possible.length);
    rolled.push(possible.splice(index, 1)[0]);
  }

  return rolled;
}

function createCraftedMap(profile) {
  const biomes = getUnlockedCraftableBiomes(profile);
  const biome = biomes[Math.floor(Math.random() * biomes.length)];
  const durationHours = randomIntBetween(5, 10);
  const createdAt = Date.now();
  const expiresAt = createdAt + durationHours * 60 * 60 * 1000;
  const rolledDangerModifiers = rollCraftedMapDangerModifiers();

  const dangerModifiers = {
    enemySpeed: 1,
    enemyHp: 1,
    bossInterval: DEFAULT_BOSS_INTERVAL,
  };

  rolledDangerModifiers.forEach((modifier) => {
    dangerModifiers[modifier.key] = modifier.value;
  });

  const rewardMultipliers = {
    gold: randomBetween(1.25, 2.5),
    xp: randomBetween(1.25, 2.5),
    rareDrop: randomBetween(1.1, 1.75),
    modifierRarity: randomBetween(1.05, 1.3),
  };

  const sourceMapId =
    biome === "Ember Ridge"
      ? "ember_ridge"
      : biome === "Tidelands"
        ? "tidelands"
        : "grasslands";

  return {
    id: `crafted_map_${createdAt}_${Math.floor(Math.random() * 100000)}`,
    name: `${biome} Expedition`,
    biome,
    durationHours,
    createdAt,
    expiresAt,
    activatedAt: null,
    slotIndex: null,
    isActive: false,
    sourceMapId,
    modifiers: [],
    rewardMultipliers,
    dangerModifiers,
    rolledDangerModifiers,
  };
}

function createNormalMapSlotEntry(mapId) {
  return { kind: "normal", mapId };
}

function createCraftedMapSlotEntry(craftedMapId) {
  return { kind: "crafted", craftedMapId };
}

function isCraftedMapSlotEntry(slotEntry) {
  return Boolean(slotEntry && typeof slotEntry === "object" && slotEntry.kind === "crafted");
}

function createInitialMapSlots() {
  return [null, null, null];
}

function getCraftedRuntimeMapId(craftedMapId) {
  return `crafted:${craftedMapId}`;
}

function getCraftedMapById(profile, craftedMapId) {
  return (profile?.craftedMaps || []).find((item) => item.id === craftedMapId) || null;
}

function isCraftedMapExpired(craftedMap) {
  if (!craftedMap?.expiresAt) return false;
  return Date.now() >= craftedMap.expiresAt;
}

function getCraftedMapTimeRemainingMs(craftedMap) {
  return Math.max(0, Number(craftedMap?.expiresAt || 0) - Date.now());
}

function buildCraftedMapPreview(craftedMap) {
  const rewardMultipliers = craftedMap?.rewardMultipliers || {};
  const rolledDangerModifiers = Array.isArray(craftedMap?.rolledDangerModifiers)
    ? craftedMap.rolledDangerModifiers
    : [];

  return {
    craftedMapId: craftedMap?.id || null,
    name: craftedMap?.name || "Crafted Map",
    biome: craftedMap?.biome || "Grasslands",
    isActive: Boolean(craftedMap?.isActive),
    isExpired: isCraftedMapExpired(craftedMap),
    durationHours: Number(craftedMap?.durationHours || 0),
    timeRemainingMs: getCraftedMapTimeRemainingMs(craftedMap),
    rewardSummary: {
      gold: Number(rewardMultipliers.gold ?? 1),
      xp: Number(rewardMultipliers.xp ?? 1),
      rareDrop: Number(rewardMultipliers.rareDrop ?? 1),
      modifierRarity: Number(rewardMultipliers.modifierRarity ?? 1),
    },
    dangerSummary: rolledDangerModifiers.map((modifier) => ({
      key: modifier?.key || "",
      label: modifier?.label || modifier?.key || "Modifier",
      value: modifier?.value ?? null,
      isBossInterval: modifier?.key === "bossInterval",
    })),
  };
}

async function handleAdminProfileAction(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const action = String(body?.action || "");
  const amount = Math.max(0, Math.floor(Number(body?.amount || 0)));
  const material = String(body?.material || "");

  if (!accountId || !action) {
    return json({ ok: false, error: "accountId and action are required." }, { status: 400 });
  }

  const loaded = await loadFullGameplayContainers(env, accountId);
  if (loaded.error) return loaded.error;

  const { profile, mapSlots, activeRuns, snapshots } = loaded;
  const now = Date.now();

  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
    },
    unlocks: {
      ...(profile.unlocks || {}),
      maps: Array.isArray(profile?.unlocks?.maps) ? [...profile.unlocks.maps] : ["grasslands"],
      towers: Array.isArray(profile?.unlocks?.towers) ? [...profile.unlocks.towers] : ["archer", "cannon"],
    },
    stats: {
      ...(profile.stats || {}),
    },
  };

  switch (action) {
    case "add_gold":
      nextProfile.gold = Number(nextProfile.gold || 0) + amount;
      break;

    case "add_xp":
      nextProfile.xp = Number(nextProfile.xp || 0) + amount;
      break;

    case "add_map_fragments":
      nextProfile.inventory.mapFragments = Number(nextProfile.inventory.mapFragments || 0) + amount;
      break;

    case "add_echo_shards":
      nextProfile.inventory.echoShards = Number(nextProfile.inventory.echoShards || 0) + amount;
      break;

    case "add_material": {
      const materialName = material || "Sunpetal";
      nextProfile.inventory[materialName] = Number(nextProfile.inventory[materialName] || 0) + amount;
      break;
    }

    case "unlock_all_maps":
      nextProfile.unlocks.maps = SAMPLE_MAPS.map((mapItem) => mapItem.id);
      break;

    case "unlock_all_towers":
      nextProfile.unlocks.towers = Object.keys(TOWER_TYPES);
      break;

    case "max_map_slots":
      nextProfile.unlocks.maxConcurrentMaps = 3;
      break;

    case "unlock_all":
      nextProfile.unlocks.maps = SAMPLE_MAPS.map((mapItem) => mapItem.id);
      nextProfile.unlocks.towers = Object.keys(TOWER_TYPES);
      nextProfile.unlocks.maxConcurrentMaps = 3;
      break;

    default:
      return json({ ok: false, error: `Unknown admin action: ${action}` }, { status: 400 });
  }

  await persistGameplayContainers(env, accountId, nextProfile, mapSlots, activeRuns, snapshots, now);

  return json({
    ok: true,
    action,
    profile: nextProfile,
    updatedAt: now,
  });
}

async function handleGetCraftedMapPreview(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");
  const craftedMapId = url.searchParams.get("craftedMapId");

  if (!accountId || !craftedMapId) {
    return json({ ok: false, error: "accountId and craftedMapId are required." }, { status: 400 });
  }

  const loaded = await loadFullGameplayContainers(env, accountId);
  if (loaded.error) return loaded.error;

  const { profile } = loaded;
  const craftedMap = getCraftedMapById(profile, craftedMapId);

  if (!craftedMap) {
    return json({ ok: false, error: "Crafted map not found." }, { status: 404 });
  }

  return json({
    ok: true,
    preview: buildCraftedMapPreview(craftedMap),
    serverTime: Date.now(),
  });
}

async function loadFullGameplayContainers(env, accountId) {
  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return { error: json({ ok: false, error: "Profile not found." }, { status: 404 }) };
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return { error: json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 }) };
  }

  const mapSlotsRow = await env.DB
    .prepare("SELECT map_slots_json FROM map_slots WHERE account_id = ?")
    .bind(accountId)
    .first();

  let mapSlots = createInitialMapSlots();
  try {
    const parsed = JSON.parse(mapSlotsRow?.map_slots_json || "[null,null,null]");
    if (Array.isArray(parsed)) {
      mapSlots = parsed.slice(0, 3);
      while (mapSlots.length < 3) mapSlots.push(null);
    }
  } catch {
    mapSlots = createInitialMapSlots();
  }

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  const snapshotsRows = await env.DB
    .prepare("SELECT run_key, snapshot_json FROM run_snapshots WHERE account_id = ?")
    .bind(accountId)
    .all();

  const activeRuns = (activeRunsRows?.results || []).map((row) => ({
    slotIndex: row.slot_index,
    runtimeKey: row.runtime_key,
    mapId: row.map_id,
    sourceMapId: row.source_map_id,
    isCrafted: Boolean(row.is_crafted),
    craftedMapId: row.crafted_map_id,
    status: row.status,
    startedAt: row.started_at,
    lastSimulatedAt: row.last_simulated_at,
    selected: Boolean(row.selected),
    opened: Boolean(row.opened),
    wave: row.wave,
    highestWaveReached: row.highest_wave_reached,
    updatedAt: row.updated_at,
  }));

  const snapshots = new Map();
  for (const row of snapshotsRows?.results || []) {
    try {
      snapshots.set(row.run_key, JSON.parse(row.snapshot_json));
    } catch {
      // skip invalid snapshot row
    }
  }

  return { profile, mapSlots, activeRuns, snapshots };
}

async function persistGameplayContainers(env, accountId, profile, mapSlots, activeRuns, snapshotsByRunKey, now) {
  const normalizedMapSlots = (Array.isArray(mapSlots) ? mapSlots.slice(0, 3) : createInitialMapSlots());
  while (normalizedMapSlots.length < 3) normalizedMapSlots.push(null);

  const statements = [
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(profile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO map_slots (account_id, map_slots_json, updated_at)
         VALUES (?, ?, ?)
         ON CONFLICT(account_id) DO UPDATE SET
           map_slots_json = excluded.map_slots_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, JSON.stringify(normalizedMapSlots), now),
    env.DB.prepare("DELETE FROM active_runs WHERE account_id = ?").bind(accountId),
  ];

  for (const item of activeRuns || []) {
    statements.push(
      env.DB
        .prepare(
          `INSERT INTO active_runs (
            account_id, slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
            status, started_at, last_simulated_at, selected, opened, wave, highest_wave_reached, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .bind(
          accountId,
          Number(item?.slotIndex ?? 0),
          item?.runtimeKey ?? null,
          item?.mapId ?? null,
          item?.sourceMapId ?? null,
          item?.isCrafted ? 1 : 0,
          item?.craftedMapId ?? null,
          item?.status ?? "running",
          item?.startedAt ?? null,
          item?.lastSimulatedAt ?? now,
          item?.selected ? 1 : 0,
          item?.opened ? 1 : 0,
          Number(item?.wave ?? 1),
          Number(item?.highestWaveReached ?? 1),
          now
        )
    );
  }

  const snapshotKeys = Array.from((snapshotsByRunKey || new Map()).keys());
  for (const runKey of snapshotKeys) {
    const snapshot = snapshotsByRunKey.get(runKey);
    statements.push(
      env.DB
        .prepare(
          `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(account_id, run_key) DO UPDATE SET
             snapshot_json = excluded.snapshot_json,
             updated_at = excluded.updated_at`
        )
        .bind(accountId, runKey, JSON.stringify(snapshot), now)
    );
  }

  await env.DB.batch(statements);
}

async function handleCraftMap(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const loaded = await loadFullGameplayContainers(env, accountId);
  if (loaded.error) return loaded.error;

  const { profile, mapSlots, activeRuns, snapshots } = loaded;
  const fragments = Number(profile?.inventory?.mapFragments || 0);
  if (fragments < MAP_CRAFT_FRAGMENT_COST) {
    return json({ ok: false, error: "Not enough Map Fragments." }, { status: 400 });
  }

  const craftedMap = createCraftedMap(profile);
  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
      mapFragments: fragments - MAP_CRAFT_FRAGMENT_COST,
    },
    craftedMaps: [craftedMap, ...(Array.isArray(profile.craftedMaps) ? profile.craftedMaps : [])],
  };

  const now = Date.now();
  await persistGameplayContainers(env, accountId, nextProfile, mapSlots, activeRuns, snapshots, now);

  return json({
    ok: true,
    craftedMap,
    profile: nextProfile,
    updatedAt: now,
  });
}

async function handleAssignMapSlot(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const slotIndex = Number(body?.slotIndex);
  const mapId = String(body?.mapId || "");

  if (!accountId || !Number.isInteger(slotIndex) || slotIndex < 0 || slotIndex > 2 || !mapId) {
    return json({ ok: false, error: "accountId, slotIndex, and mapId are required." }, { status: 400 });
  }

  const loaded = await loadFullGameplayContainers(env, accountId);
  if (loaded.error) return loaded.error;

  const { profile, mapSlots, activeRuns, snapshots } = loaded;
  const maxSlots = Math.max(1, Number(profile?.unlocks?.maxConcurrentMaps || 1));
  if (slotIndex >= maxSlots) {
    return json({ ok: false, error: "That map slot is locked." }, { status: 400 });
  }

  const baseMap = SAMPLE_MAPS.find((item) => item.id === mapId);
  if (!baseMap) {
    return json({ ok: false, error: "Map not found." }, { status: 404 });
  }

  const nextMapSlots = [...mapSlots];
  while (nextMapSlots.length < 3) nextMapSlots.push(null);
  nextMapSlots[slotIndex] = createNormalMapSlotEntry(mapId);

  const now = Date.now();
  const existingRunIndex = activeRuns.findIndex((item) => Number(item.slotIndex) === slotIndex);
  const runtimeKey = mapId;

  if (existingRunIndex >= 0) {
    activeRuns[existingRunIndex] = {
      ...activeRuns[existingRunIndex],
      slotIndex,
      runtimeKey,
      mapId,
      sourceMapId: mapId,
      isCrafted: false,
      craftedMapId: null,
      opened: false,
      selected: true,
      updatedAt: now,
      lastSimulatedAt: activeRuns[existingRunIndex].lastSimulatedAt || now,
    };
  } else {
    activeRuns.push({
      slotIndex,
      runtimeKey,
      mapId,
      sourceMapId: mapId,
      isCrafted: false,
      craftedMapId: null,
      status: "running",
      startedAt: now,
      lastSimulatedAt: now,
      selected: true,
      opened: false,
      wave: 1,
      highestWaveReached: 1,
      updatedAt: now,
    });
  }

  for (const item of activeRuns) {
    if (Number(item.slotIndex) !== slotIndex) item.selected = false;
  }

  if (!snapshots.has(runtimeKey)) {
    snapshots.set(runtimeKey, createFallbackSnapshotFromActiveRun({
      slotIndex,
      runtimeKey,
      mapId,
      sourceMapId: mapId,
      status: "running",
      startedAt: now,
      lastSimulatedAt: now,
      wave: 1,
      highestWaveReached: 1,
    }, baseMap, now));
  }

  const nextProfile = {
    ...profile,
    appState: {
      ...(profile.appState || {}),
      activeMapId: mapId,
      selectedMapId: mapId,
    },
  };

  await persistGameplayContainers(env, accountId, nextProfile, nextMapSlots, activeRuns, snapshots, now);

  return json({
    ok: true,
    mapSlots: nextMapSlots,
    profile: nextProfile,
    activeRuns,
    updatedAt: now,
  });
}

async function handleOpenRunSession(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  const activeRuns = (activeRunsRows?.results || []).map((row) => ({
    slotIndex: row.slot_index,
    runtimeKey: row.runtime_key,
    mapId: row.map_id,
    sourceMapId: row.source_map_id,
    isCrafted: Boolean(row.is_crafted),
    craftedMapId: row.crafted_map_id,
    status: row.status,
    startedAt: row.started_at,
    lastSimulatedAt: row.last_simulated_at,
    selected: Boolean(row.selected),
    opened: Boolean(row.opened),
    wave: row.wave,
    highestWaveReached: row.highest_wave_reached,
    updatedAt: row.updated_at,
  }));

  const target = activeRuns.find((item) => item.runtimeKey === runtimeKey);
  if (!target) {
    return json({ ok: false, error: "Run not found." }, { status: 404 });
  }

  const now = Date.now();
  const statements = [];

  activeRuns.forEach((item) => {
    statements.push(
      env.DB
        .prepare(
          `UPDATE active_runs
           SET opened = ?, selected = ?, updated_at = ?
           WHERE account_id = ? AND runtime_key = ?`
        )
        .bind(
          item.runtimeKey === runtimeKey ? 1 : 0,
          item.runtimeKey === runtimeKey ? 1 : 0,
          now,
          accountId,
          item.runtimeKey
        )
    );
  });

  await env.DB.batch(statements);

  return json({
    ok: true,
    runtimeKey,
    updatedAt: now,
  });
}

async function handleCloseRunSession(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const row = await env.DB
    .prepare(
      `SELECT runtime_key FROM active_runs
       WHERE account_id = ? AND runtime_key = ?`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!row?.runtime_key) {
    return json({ ok: false, error: "Run not found." }, { status: 404 });
  }

  const now = Date.now();
  await env.DB
    .prepare(
      `UPDATE active_runs
       SET opened = 0, updated_at = ?
       WHERE account_id = ? AND runtime_key = ?`
    )
    .bind(now, accountId, runtimeKey)
    .run();

  return json({
    ok: true,
    runtimeKey,
    updatedAt: now,
  });
}

async function handleSelectRunSession(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT runtime_key
       FROM active_runs
       WHERE account_id = ?`
    )
    .bind(accountId)
    .all();

  const runtimeKeys = (activeRunsRows?.results || []).map((row) => row.runtime_key);
  if (!runtimeKeys.includes(runtimeKey)) {
    return json({ ok: false, error: "Run not found." }, { status: 404 });
  }

  const now = Date.now();
  const statements = runtimeKeys.map((key) =>
    env.DB
      .prepare(
        `UPDATE active_runs
         SET selected = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(key === runtimeKey ? 1 : 0, now, accountId, key)
  );

  await env.DB.batch(statements);

  return json({
    ok: true,
    runtimeKey,
    updatedAt: now,
  });
}

async function handleAssignCraftedMapSlot(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const slotIndex = Number(body?.slotIndex);
  const craftedMapId = String(body?.craftedMapId || "");

  if (!accountId || !Number.isInteger(slotIndex) || slotIndex < 0 || slotIndex > 2 || !craftedMapId) {
    return json({ ok: false, error: "accountId, slotIndex, and craftedMapId are required." }, { status: 400 });
  }

  const loaded = await loadFullGameplayContainers(env, accountId);
  if (loaded.error) return loaded.error;

  const { profile, mapSlots, activeRuns, snapshots } = loaded;
  const maxSlots = Math.max(1, Number(profile?.unlocks?.maxConcurrentMaps || 1));
  if (slotIndex >= maxSlots) {
    return json({ ok: false, error: "That map slot is locked." }, { status: 400 });
  }

  const craftedMap = getCraftedMapById(profile, craftedMapId);
  if (!craftedMap) {
    return json({ ok: false, error: "Crafted map not found." }, { status: 404 });
  }
  if (isCraftedMapExpired(craftedMap)) {
    return json({ ok: false, error: "Crafted map has expired." }, { status: 400 });
  }

  const runtimeKey = getCraftedRuntimeMapId(craftedMapId);
  const nextMapSlots = [...mapSlots];
  while (nextMapSlots.length < 3) nextMapSlots.push(null);
  nextMapSlots[slotIndex] = createCraftedMapSlotEntry(craftedMapId);

  const now = Date.now();
  const existingRunIndex = activeRuns.findIndex((item) => Number(item.slotIndex) === slotIndex);

  if (existingRunIndex >= 0) {
    activeRuns[existingRunIndex] = {
      ...activeRuns[existingRunIndex],
      slotIndex,
      runtimeKey,
      mapId: runtimeKey,
      sourceMapId: craftedMap.sourceMapId || "grasslands",
      isCrafted: true,
      craftedMapId,
      opened: false,
      selected: true,
      updatedAt: now,
      lastSimulatedAt: activeRuns[existingRunIndex].lastSimulatedAt || now,
    };
  } else {
    activeRuns.push({
      slotIndex,
      runtimeKey,
      mapId: runtimeKey,
      sourceMapId: craftedMap.sourceMapId || "grasslands",
      isCrafted: true,
      craftedMapId,
      status: "running",
      startedAt: now,
      lastSimulatedAt: now,
      selected: true,
      opened: false,
      wave: 1,
      highestWaveReached: 1,
      updatedAt: now,
    });
  }

  for (const item of activeRuns) {
    if (Number(item.slotIndex) != slotIndex) item.selected = false;
  }

  const nextCraftedMaps = (profile.craftedMaps || []).map((item) =>
    item.id === craftedMapId
      ? {
          ...item,
          isActive: true,
          activatedAt: item.activatedAt || now,
          slotIndex,
        }
      : item
  );

  const nextProfile = {
    ...profile,
    craftedMaps: nextCraftedMaps,
    appState: {
      ...(profile.appState || {}),
      activeMapId: craftedMap.sourceMapId || "grasslands",
      selectedMapId: craftedMap.sourceMapId || "grasslands",
    },
  };

  if (!snapshots.has(runtimeKey)) {
    const baseMap = SAMPLE_MAPS.find((item) => item.id === (craftedMap.sourceMapId || "grasslands")) || SAMPLE_MAPS[0];
    snapshots.set(runtimeKey, createFallbackSnapshotFromActiveRun({
      slotIndex,
      runtimeKey,
      mapId: runtimeKey,
      sourceMapId: craftedMap.sourceMapId || baseMap.id,
      isCrafted: true,
      craftedMapId,
      status: "running",
      startedAt: now,
      lastSimulatedAt: now,
      wave: 1,
      highestWaveReached: 1,
    }, {
      ...baseMap,
      id: runtimeKey,
      craftedMapId,
      craftedData: craftedMap,
      isCraftedMap: true,
      name: craftedMap.name,
      biome: craftedMap.biome,
    }, now));
  }

  await persistGameplayContainers(env, accountId, nextProfile, nextMapSlots, activeRuns, snapshots, now);

  return json({
    ok: true,
    mapSlots: nextMapSlots,
    profile: nextProfile,
    activeRuns,
    updatedAt: now,
  });
}

function createInitialProfile({ accountId, username, createdAt }) {
  return {
    profileVersion: PROFILE_SCHEMA_VERSION,
    id: accountId,
    name: username,
    hasCompletedSignup: true,
    createdAt,
    level: 1,
    xp: 0,
    gold: 30,
    inventory: {
      Sunpetal: 0,
      "Ember Shard": 0,
      Tideglass: 0,
      mapFragments: 0,
      echoShards: 0
    },
    modifiers: [],
    craftedMaps: [],
    unlocks: {
      maps: ["grasslands"],
      towers: ["archer", "cannon"],
      maxConcurrentMaps: 1
    },
    permanentTowerUpgrades: {
      archer: { damage: 0, range: 0 },
      cannon: { damage: 0, range: 0 }
    },
    towerLevels: {
      archer: { xp: 0 },
      cannon: { xp: 0 },
      frost: { xp: 0 }
    },
    towerMasteryLoadouts: {
      archer: [],
      cannon: [],
      frost: []
    },
    towerMasteryDrafts: {
      archer: { selectedNodeIds: [], purchasedNodeIds: [] },
      cannon: { selectedNodeIds: [], purchasedNodeIds: [] },
      frost: { selectedNodeIds: [], purchasedNodeIds: [] }
    },
    biomeProgress: {
      grasslands: { points: 0, dropBonusPct: 0 },
      ember: { points: 0, dropBonusPct: 0 },
      tide: { points: 0, dropBonusPct: 0 }
    },
    stats: {
      lifetimeGoldEarned: 0,
      highestWaveReached: 1,
      mapsCompleted: 0
    },
    admin: {
      godMode: false,
      simulationSpeed: 1
    }
  };
}

async function readJson(request) {
  try {
    return await request.json();
  } catch {
    return null;
  }
}

function normalizeChatMessageRow(row) {
  return {
    id: row.id,
    channel: "global",
    messageType: "player",
    author: row.author,
    text: row.message_text,
    createdAt: row.created_at,
    accountId: row.account_id,
  };
}

async function handleGetGlobalChat(request, env) {
  const url = new URL(request.url);
  const limitParam = Number(url.searchParams.get("limit") || 50);
  const limit = Math.max(1, Math.min(100, Number.isFinite(limitParam) ? limitParam : 50));

  const rows = await env.DB
    .prepare(
      `SELECT id, account_id, author, message_text, created_at
       FROM global_chat_messages
       ORDER BY created_at DESC
       LIMIT ?`
    )
    .bind(limit)
    .all();

  const messages = (rows?.results || [])
    .map(normalizeChatMessageRow)
    .sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));

  return json({
    ok: true,
    messages,
  });
}

async function handleSendGlobalChat(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const text = String(body?.text || "").trim();

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  if (!text) {
    return json({ ok: false, error: "Message text is required." }, { status: 400 });
  }

  if (text.length > GLOBAL_CHAT_MAX_LENGTH) {
    return json({ ok: false, error: `Message must be ${GLOBAL_CHAT_MAX_LENGTH} characters or fewer.` }, { status: 400 });
  }

  const accountRow = await env.DB
    .prepare("SELECT id, username FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!accountRow) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const createdAt = Date.now();

  const recentOwnMessage = await env.DB
    .prepare(
      `SELECT id, message_text, created_at
       FROM global_chat_messages
       WHERE account_id = ?
       ORDER BY created_at DESC
       LIMIT 1`
    )
    .bind(accountId)
    .first();

  if (recentOwnMessage?.created_at) {
    const msSinceLastMessage = createdAt - Number(recentOwnMessage.created_at || 0);
    if (msSinceLastMessage < GLOBAL_CHAT_SEND_COOLDOWN_MS) {
      return json(
        {
          ok: false,
          error: "You are sending messages too quickly.",
          retryAfterMs: GLOBAL_CHAT_SEND_COOLDOWN_MS - msSinceLastMessage,
          errorCode: "CHAT_RATE_LIMIT",
        },
        { status: 429 }
      );
    }
  }

  const recentDuplicate = await env.DB
    .prepare(
      `SELECT id, created_at
       FROM global_chat_messages
       WHERE account_id = ? AND message_text = ?
       ORDER BY created_at DESC
       LIMIT 1`
    )
    .bind(accountId, text)
    .first();

  if (recentDuplicate?.created_at) {
    const msSinceDuplicate = createdAt - Number(recentDuplicate.created_at || 0);
    if (msSinceDuplicate < GLOBAL_CHAT_DUPLICATE_WINDOW_MS) {
      return json(
        {
          ok: false,
          error: "You just sent that exact message.",
          retryAfterMs: GLOBAL_CHAT_DUPLICATE_WINDOW_MS - msSinceDuplicate,
          errorCode: "CHAT_DUPLICATE",
        },
        { status: 429 }
      );
    }
  }

  const messageId = generateId("chat");

  await env.DB.batch([
    env.DB
      .prepare(
        `INSERT INTO global_chat_messages (id, account_id, author, message_text, created_at)
         VALUES (?, ?, ?, ?, ?)`
      )
      .bind(messageId, accountId, accountRow.username, text, createdAt),
    env.DB
      .prepare(
        `DELETE FROM global_chat_messages
         WHERE id IN (
           SELECT id FROM global_chat_messages
           ORDER BY created_at DESC
           LIMIT -1 OFFSET 200
         )`
      )
      .bind(),
  ]);

  return json({
    ok: true,
    message: {
      id: messageId,
      channel: "global",
      messageType: "player",
      author: accountRow.username,
      text,
      createdAt,
      accountId,
    },
  });
}

async function handleSignup(request, env) {
  const body = await readJson(request);
  const username = body?.username ?? "";
  const validationError = validateUsername(username);

  if (validationError) {
    return json({ ok: false, error: validationError }, { status: 400 });
  }

  const trimmedUsername = username.trim();
  const createdAt = Date.now();
  const accountId = generateId("acct");

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE username = ?")
    .bind(trimmedUsername)
    .first();

  if (existing) {
    return json({ ok: false, error: "That username is already taken." }, { status: 409 });
  }

  const profile = createInitialProfile({
    accountId,
    username: trimmedUsername,
    createdAt,
  });

  await env.DB.batch([
    env.DB
      .prepare("INSERT INTO accounts (id, username, created_at) VALUES (?, ?, ?)")
      .bind(accountId, trimmedUsername, createdAt),
    env.DB
      .prepare("INSERT INTO player_profiles (account_id, profile_json, updated_at) VALUES (?, ?, ?)")
      .bind(accountId, JSON.stringify(profile), createdAt),
  ]);

  return json({
    ok: true,
    accountId,
    username: trimmedUsername,
    profile,
  });
}

async function handleGetProfile(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const row = await env.DB
    .prepare("SELECT profile_json, updated_at FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!row) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile = null;
  try {
    profile = JSON.parse(row.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  return json({
    ok: true,
    profile,
    updatedAt: row.updated_at,
  });
}

async function handleSaveProfile(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const profile = body?.profile;

  if (!accountId || !profile || typeof profile !== "object") {
    return json({ ok: false, error: "accountId and profile are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT account_id FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
    .bind(JSON.stringify(profile), updatedAt, accountId)
    .run();

  return json({
    ok: true,
    updatedAt,
  });
}


async function runBackendCatchupForAccount(env, accountId) {
  if (!accountId) {
    return { ok: false, error: "accountId is required." };
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return { ok: false, error: "Profile not found." };
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return { ok: false, error: "Stored profile is invalid JSON." };
  }

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  const snapshotRows = await env.DB
    .prepare("SELECT run_key, snapshot_json, updated_at FROM run_snapshots WHERE account_id = ?")
    .bind(accountId)
    .all();

  const snapshotMap = new Map();
  for (const row of snapshotRows?.results || []) {
    try {
      snapshotMap.set(row.run_key, JSON.parse(row.snapshot_json));
    } catch {
      // skip bad snapshot
    }
  }

  const now = Date.now();
  let totalGoldEarned = 0;
  let totalPlayerXpEarned = 0;
  let totalMapFragmentsEarned = 0;
  let changedRuns = 0;
  const debugRuns = [];
  const notificationsToAdd = [];
  const snapshotStatements = [];
  const activeRunStatements = [];

  const nextInventory = { ...(profile.inventory || {}) };
  const nextBiomeProgress = { ...(profile.biomeProgress || {}) };

  for (const row of activeRunsRows?.results || []) {
    if (row.status !== "running") continue;

    const activeRun = {
      slotIndex: row.slot_index,
      runtimeKey: row.runtime_key,
      mapId: row.map_id,
      sourceMapId: row.source_map_id,
      isCrafted: Boolean(row.is_crafted),
      craftedMapId: row.crafted_map_id,
      status: row.status,
      startedAt: row.started_at,
      lastSimulatedAt: row.last_simulated_at,
      selected: Boolean(row.selected),
      opened: Boolean(row.opened),
      wave: row.wave,
      highestWaveReached: row.highest_wave_reached,
      updatedAt: row.updated_at,
    };

    const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
    if (!runtimeMap) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        skipped: "missing_runtime_map",
      });
      continue;
    }

    const hadStoredSnapshot = snapshotMap.has(activeRun.runtimeKey);
    const snapshot =
      snapshotMap.get(activeRun.runtimeKey) ||
      createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, now);

    const snapshotLastSimulatedAt = Number(snapshot.lastSimulatedAt || 0);
    const activeRunLastSimulatedAt = Number(activeRun.lastSimulatedAt || 0);

    const safestLastSimulatedAt =
      snapshotLastSimulatedAt > 0
        ? snapshotLastSimulatedAt
        : activeRunLastSimulatedAt;

    if (!safestLastSimulatedAt) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        hadStoredSnapshot,
        skipped: "missing_last_simulated_at",
      });
      continue;
    }

    const missedMs = Math.max(0, now - safestLastSimulatedAt);
    if (missedMs <= 0) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        hadStoredSnapshot,
        safestLastSimulatedAt,
        now,
        missedMs,
        skipped: "non_positive_missed_ms",
      });
      continue;
    }

    const enemyIdStart =
      Math.max(1, ...((snapshot.enemies || []).map((enemy) => Number(enemy.id) || 0)), 0) + 1;

    const waveBefore = Number(snapshot.wave || activeRun.wave || 1);
    const enemiesBefore = Array.isArray(snapshot.enemies) ? snapshot.enemies.length : 0;
    const enemyPathPositionsBefore = Array.isArray(snapshot.enemies)
      ? snapshot.enemies.slice(0, 5).map((enemy) => Number(enemy?.pathPosition || 0))
      : [];

    const catchupResult = applyOfflineCatchupToRun({
      mapData: runtimeMap,
      runState: {
        ...snapshot,
        isRunning: true,
      },
      missedMs,
      playerProfile: profile,
      enemyIdStart,
    });

    const enemyPathPositionsAfter = Array.isArray(catchupResult?.nextRunState?.enemies)
      ? catchupResult.nextRunState.enemies.slice(0, 5).map((enemy) => Number(enemy?.pathPosition || 0))
      : [];

    debugRuns.push({
      runtimeKey: activeRun.runtimeKey,
      hadStoredSnapshot,
      safestLastSimulatedAt,
      now,
      missedMs,
      appliedMs: catchupResult.appliedMs || 0,
      waveBefore,
      waveAfter: Number(catchupResult?.nextRunState?.wave || waveBefore),
      enemiesBefore,
      enemiesAfter: Array.isArray(catchupResult?.nextRunState?.enemies)
        ? catchupResult.nextRunState.enemies.length
        : enemiesBefore,
      enemyPathPositionsBefore,
      enemyPathPositionsAfter,
      goldEarned: Number(catchupResult?.goldEarned || 0),
      xpEarned: Number(catchupResult?.playerXpEarned || 0),
      changed:
        Boolean(catchupResult.appliedMs) &&
        (
          Number(catchupResult?.nextRunState?.wave || waveBefore) !== waveBefore ||
          (Array.isArray(catchupResult?.nextRunState?.enemies)
            ? catchupResult.nextRunState.enemies.length
            : enemiesBefore) !== enemiesBefore ||
          JSON.stringify(enemyPathPositionsBefore) !== JSON.stringify(enemyPathPositionsAfter) ||
          Number(catchupResult?.goldEarned || 0) > 0
        ),
    });

    if (!catchupResult.appliedMs) continue;

    changedRuns += 1;
    totalGoldEarned += catchupResult.goldEarned || 0;
    totalPlayerXpEarned += catchupResult.playerXpEarned || 0;
    totalMapFragmentsEarned += catchupResult.mapFragmentsEarned || 0;

    const rareMaterialName = getRareMaterialNameForMap(runtimeMap);
    if ((catchupResult.rareDropsEarned || 0) > 0) {
      nextInventory[rareMaterialName] =
        (nextInventory[rareMaterialName] || 0) + (catchupResult.rareDropsEarned || 0);
    }

    if ((catchupResult.mapFragmentsEarned || 0) > 0) {
      nextInventory.mapFragments =
        (nextInventory.mapFragments || 0) + (catchupResult.mapFragmentsEarned || 0);
    }

    const biomeKey = getBiomeKeyFromMap(runtimeMap);
    const currentBiome = nextBiomeProgress[biomeKey] || { points: 0, dropBonusPct: 0 };
    const nextPoints = currentBiome.points + (catchupResult.biomeProgressEarned || 0);
    nextBiomeProgress[biomeKey] = {
      points: nextPoints,
      dropBonusPct: nextPoints / 10000,
    };

    if ((catchupResult.appliedMs || 0) >= OFFLINE_NOTIFICATION_MIN_MS) {
      notificationsToAdd.push(
        createOfflineCatchupNotification({
          mapName: runtimeMap.name,
          appliedMs: catchupResult.appliedMs,
          goldEarned: catchupResult.goldEarned || 0,
          xpEarned: catchupResult.playerXpEarned || 0,
          rareDrops: catchupResult.rareDropsEarned || 0,
          mapFragments: catchupResult.mapFragmentsEarned || 0,
          modifiers: 0,
        })
      );
    }

    snapshotStatements.push(
      env.DB
        .prepare(
          `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(account_id, run_key) DO UPDATE SET
             snapshot_json = excluded.snapshot_json,
             updated_at = excluded.updated_at`
        )
        .bind(accountId, activeRun.runtimeKey, JSON.stringify(catchupResult.nextRunState), now)
    );

    activeRunStatements.push(
      env.DB
        .prepare(
          `UPDATE active_runs
           SET wave = ?, highest_wave_reached = ?, last_simulated_at = ?, updated_at = ?
           WHERE account_id = ? AND slot_index = ?`
        )
        .bind(
          Number(catchupResult.nextRunState.wave || activeRun.wave || 1),
          Number(catchupResult.nextRunState.highestWaveReached || activeRun.highestWaveReached || 1),
          now,
          now,
          accountId,
          Number(activeRun.slotIndex)
        )
    );
  }

  if (!changedRuns) {
    return {
      ok: true,
      changedRuns: 0,
      totalGoldEarned: 0,
      totalPlayerXpEarned: 0,
      totalMapFragmentsEarned: 0,
      notificationsAdded: 0,
      debugRuns,
    };
  }

  const nextProfile = {
    ...profile,
    gold: (profile.gold || 0) + totalGoldEarned,
    inventory: nextInventory,
    biomeProgress: nextBiomeProgress,
    stats: {
      ...(profile.stats || {}),
      lifetimeGoldEarned: (profile.stats?.lifetimeGoldEarned || 0) + totalGoldEarned,
      highestWaveReached: profile.stats?.highestWaveReached || 1,
    },
  };

  const notificationsRow = await env.DB
    .prepare("SELECT notifications_json FROM notifications WHERE account_id = ?")
    .bind(accountId)
    .first();

  let existingNotifications = [];
  try {
    existingNotifications = JSON.parse(notificationsRow?.notifications_json || "[]");
    if (!Array.isArray(existingNotifications)) existingNotifications = [];
  } catch {
    existingNotifications = [];
  }

  const mergedNotifications = [...notificationsToAdd, ...existingNotifications]
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0))
    .slice(0, 25);

  const highestWaveReachedAfterCatchup = Math.max(
    profile.stats?.highestWaveReached || 1,
    ...((activeRunsRows?.results || []).map((row) => Number(row.highest_wave_reached || row.wave || 1)))
  );

  nextProfile.stats.highestWaveReached = highestWaveReachedAfterCatchup;

  const statements = [
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO notifications (account_id, notifications_json, updated_at)
         VALUES (?, ?, ?)
         ON CONFLICT(account_id) DO UPDATE SET
           notifications_json = excluded.notifications_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, JSON.stringify(mergedNotifications), now),
    ...snapshotStatements,
    ...activeRunStatements,
  ];

  await env.DB.batch(statements);

  return {
    ok: true,
    changedRuns,
    totalGoldEarned,
    totalPlayerXpEarned,
    totalMapFragmentsEarned,
    notificationsAdded: notificationsToAdd.length,
    debugRuns,
  };
}

async function handleGetRunSnapshots(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  await runBackendCatchupForAccount(env, accountId);

  const rows = await env.DB
    .prepare("SELECT run_key, snapshot_json, updated_at FROM run_snapshots WHERE account_id = ?")
    .bind(accountId)
    .all();

  const snapshots = {};

  for (const row of rows?.results || []) {
    try {
      snapshots[row.run_key] = JSON.parse(row.snapshot_json);
    } catch {
      // skip invalid row
    }
  }

  return json({
    ok: true,
    snapshots,
  });
}

async function handleSaveRunSnapshots(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const snapshots = body?.snapshots;

  if (!accountId || !snapshots || typeof snapshots !== "object") {
    return json({ ok: false, error: "accountId and snapshots are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const updatedAt = Date.now();
  const statements = [];

  const snapshotKeys = Object.keys(snapshots || {});
  for (const runKey of snapshotKeys) {
    statements.push(
      env.DB
        .prepare(
          `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(account_id, run_key) DO UPDATE SET
             snapshot_json = excluded.snapshot_json,
             updated_at = excluded.updated_at`
        )
        .bind(accountId, runKey, JSON.stringify(snapshots[runKey]), updatedAt)
    );
  }

  const existingRows = await env.DB
    .prepare("SELECT run_key FROM run_snapshots WHERE account_id = ?")
    .bind(accountId)
    .all();

  const existingKeys = new Set((existingRows?.results || []).map((row) => row.run_key));
  for (const existingKey of existingKeys) {
    if (!snapshotKeys.includes(existingKey)) {
      statements.push(
        env.DB
          .prepare("DELETE FROM run_snapshots WHERE account_id = ? AND run_key = ?")
          .bind(accountId, existingKey)
      );
    }
  }

  if (statements.length) {
    await env.DB.batch(statements);
  }

  return json({
    ok: true,
    updatedAt,
    count: snapshotKeys.length,
  });
}

async function handleGetNotifications(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const row = await env.DB
    .prepare("SELECT notifications_json, updated_at FROM notifications WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!row) {
    return json({
      ok: true,
      notifications: [],
      updatedAt: null,
    });
  }

  let notifications = [];
  try {
    notifications = JSON.parse(row.notifications_json || "[]");
  } catch {
    notifications = [];
  }

  return json({
    ok: true,
    notifications,
    updatedAt: row.updated_at,
  });
}

async function handleSaveNotifications(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const notifications = body?.notifications;

  if (!accountId || !Array.isArray(notifications)) {
    return json({ ok: false, error: "accountId and notifications are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `INSERT INTO notifications (account_id, notifications_json, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(account_id) DO UPDATE SET
         notifications_json = excluded.notifications_json,
         updated_at = excluded.updated_at`
    )
    .bind(accountId, JSON.stringify(notifications), updatedAt)
    .run();

  return json({
    ok: true,
    updatedAt,
    count: notifications.length,
  });
}

async function handleGetMapSlots(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const row = await env.DB
    .prepare("SELECT map_slots_json, updated_at FROM map_slots WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!row) {
    return json({
      ok: true,
      mapSlots: [null, null, null],
      updatedAt: null,
    });
  }

  let mapSlots = [null, null, null];
  try {
    const parsed = JSON.parse(row.map_slots_json || "[null,null,null]");
    if (Array.isArray(parsed)) {
      mapSlots = parsed.slice(0, 3);
      while (mapSlots.length < 3) mapSlots.push(null);
    }
  } catch {
    mapSlots = [null, null, null];
  }

  return json({
    ok: true,
    mapSlots,
    updatedAt: row.updated_at,
  });
}

async function handleSaveMapSlots(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const mapSlots = body?.mapSlots;

  if (!accountId || !Array.isArray(mapSlots)) {
    return json({ ok: false, error: "accountId and mapSlots are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const normalized = mapSlots.slice(0, 3);
  while (normalized.length < 3) normalized.push(null);

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `INSERT INTO map_slots (account_id, map_slots_json, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(account_id) DO UPDATE SET
         map_slots_json = excluded.map_slots_json,
         updated_at = excluded.updated_at`
    )
    .bind(accountId, JSON.stringify(normalized), updatedAt)
    .run();

  return json({
    ok: true,
    updatedAt,
    count: normalized.filter(Boolean).length,
  });
}

async function handleGetActiveRuns(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  await runBackendCatchupForAccount(env, accountId);

  const rows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  return json({
    ok: true,
    activeRuns: (rows?.results || []).map((row) => ({
      slotIndex: row.slot_index,
      runtimeKey: row.runtime_key,
      mapId: row.map_id,
      sourceMapId: row.source_map_id,
      isCrafted: Boolean(row.is_crafted),
      craftedMapId: row.crafted_map_id,
      status: row.status,
      startedAt: row.started_at,
      lastSimulatedAt: row.last_simulated_at,
      selected: Boolean(row.selected),
      opened: Boolean(row.opened),
      wave: row.wave,
      highestWaveReached: row.highest_wave_reached,
      updatedAt: row.updated_at,
    })),
  });
}

async function handleSaveActiveRuns(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const activeRuns = body?.activeRuns;

  if (!accountId || !Array.isArray(activeRuns)) {
    return json({ ok: false, error: "accountId and activeRuns are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const updatedAt = Date.now();
  const statements = [
    env.DB.prepare("DELETE FROM active_runs WHERE account_id = ?").bind(accountId),
  ];

  for (const item of activeRuns) {
    statements.push(
      env.DB
        .prepare(
          `INSERT INTO active_runs (
            account_id,
            slot_index,
            runtime_key,
            map_id,
            source_map_id,
            is_crafted,
            crafted_map_id,
            status,
            started_at,
            last_simulated_at,
            selected,
            opened,
            wave,
            highest_wave_reached,
            updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .bind(
          accountId,
          Number(item?.slotIndex ?? 0),
          item?.runtimeKey ?? null,
          item?.mapId ?? null,
          item?.sourceMapId ?? null,
          item?.isCrafted ? 1 : 0,
          item?.craftedMapId ?? null,
          item?.status ?? "running",
          item?.startedAt ?? null,
          item?.lastSimulatedAt ?? null,
          item?.selected ? 1 : 0,
          item?.opened ? 1 : 0,
          Number(item?.wave ?? 1),
          Number(item?.highestWaveReached ?? 1),
          updatedAt
        )
    );
  }

  await env.DB.batch(statements);

  return json({
    ok: true,
    updatedAt,
    count: activeRuns.length,
  });
}

async function handleUpsertActiveRun(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const activeRun = body?.activeRun;

  if (!accountId || !activeRun || typeof activeRun !== "object") {
    return json({ ok: false, error: "accountId and activeRun are required." }, { status: 400 });
  }

  if (activeRun.slotIndex == null || activeRun.runtimeKey == null) {
    return json({ ok: false, error: "activeRun.slotIndex and activeRun.runtimeKey are required." }, { status: 400 });
  }

  const existing = await env.DB
    .prepare("SELECT id FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!existing) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `INSERT INTO active_runs (
        account_id,
        slot_index,
        runtime_key,
        map_id,
        source_map_id,
        is_crafted,
        crafted_map_id,
        status,
        started_at,
        last_simulated_at,
        selected,
        opened,
        wave,
        highest_wave_reached,
        updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(account_id, slot_index) DO UPDATE SET
        runtime_key = excluded.runtime_key,
        map_id = excluded.map_id,
        source_map_id = excluded.source_map_id,
        is_crafted = excluded.is_crafted,
        crafted_map_id = excluded.crafted_map_id,
        status = excluded.status,
        started_at = excluded.started_at,
        last_simulated_at = excluded.last_simulated_at,
        selected = excluded.selected,
        opened = excluded.opened,
        wave = excluded.wave,
        highest_wave_reached = excluded.highest_wave_reached,
        updated_at = excluded.updated_at`
    )
    .bind(
      accountId,
      Number(activeRun.slotIndex ?? 0),
      activeRun.runtimeKey ?? null,
      activeRun.mapId ?? null,
      activeRun.sourceMapId ?? null,
      activeRun.isCrafted ? 1 : 0,
      activeRun.craftedMapId ?? null,
      activeRun.status ?? "running",
      activeRun.startedAt ?? null,
      activeRun.lastSimulatedAt ?? null,
      activeRun.selected ? 1 : 0,
      activeRun.opened ? 1 : 0,
      Number(activeRun.wave ?? 1),
      Number(activeRun.highestWaveReached ?? 1),
      updatedAt
    )
    .run();

  return json({
    ok: true,
    updatedAt,
    activeRun: {
      ...activeRun,
      updatedAt,
    },
  });
}

async function getExistingActiveRunForAction(env, accountId, slotIndex) {
  return env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached
       FROM active_runs
       WHERE account_id = ? AND slot_index = ?`
    )
    .bind(accountId, Number(slotIndex))
    .first();
}

function normalizeActiveRunRow(row, updatedAtOverride = null) {
  if (!row) return null;
  return {
    slotIndex: row.slot_index,
    runtimeKey: row.runtime_key,
    mapId: row.map_id,
    sourceMapId: row.source_map_id,
    isCrafted: Boolean(row.is_crafted),
    craftedMapId: row.crafted_map_id,
    status: row.status,
    startedAt: row.started_at,
    lastSimulatedAt: row.last_simulated_at,
    selected: Boolean(row.selected),
    opened: Boolean(row.opened),
    wave: row.wave,
    highestWaveReached: row.highest_wave_reached,
    updatedAt: updatedAtOverride,
  };
}


const TOWER_TYPES = {
  archer: {
    key: "archer",
    label: "Archer",
    cost: 30,
    range: 2.25,
    damage: 8,
    cooldownMs: 700,
    color: "bg-emerald-400",
  },
  cannon: {
    key: "cannon",
    label: "Cannon",
    cost: 50,
    range: 1.75,
    damage: 15,
    cooldownMs: 1400,
    color: "bg-orange-400",
  },
  frost: {
    key: "frost",
    label: "Frost Tower",
    cost: 45,
    range: 2.0,
    damage: 4,
    cooldownMs: 950,
    color: "bg-cyan-400",
    unlockLevel: 3,
    slowAmount: 0.65,
    slowDurationMs: 1800,
  },
};

const TOWER_TIER_UNLOCK_LEVELS = { small: 5, medium: 10, big: 15 };

const TOWER_MASTERY_DEFINITIONS = {
  archer: [
    { id: "archer_small_1", size: "small", branch: "precision", name: "Sharp Tip", effect: { damage: 1 }, prerequisiteNodeIds: [] },
    { id: "archer_small_2", size: "small", branch: "precision", name: "Long Draw", effect: { range: 0.15 }, prerequisiteNodeIds: ["archer_small_1"] },
    { id: "archer_medium_1", size: "medium", branch: "precision", name: "Marksman's Form", effect: { damage: 3 }, prerequisiteNodeIds: ["archer_small_2"] },
    { id: "archer_small_3", size: "small", branch: "volley", name: "Quick Grip", effect: { attackSpeedMs: 15 }, prerequisiteNodeIds: [] },
    { id: "archer_small_4", size: "small", branch: "volley", name: "Loose String", effect: { attackSpeedMs: 15 }, prerequisiteNodeIds: ["archer_small_3"] },
    { id: "archer_medium_2", size: "medium", branch: "volley", name: "Rapid Rhythm", effect: { attackSpeedMs: 35 }, prerequisiteNodeIds: ["archer_small_4"] },
    { id: "archer_big_1", size: "big", specialist: true, branch: "precision", name: "Deadeye Specialist", effect: { damage: 6, range: 0.35 }, prerequisiteNodeIds: ["archer_medium_1"] },
    { id: "archer_big_2", size: "big", specialist: true, branch: "volley", name: "Volley Specialist", effect: { attackSpeedMs: 70, damage: 2 }, prerequisiteNodeIds: ["archer_medium_2"] },
  ],
  cannon: [
    { id: "cannon_small_1", size: "small", branch: "blast", name: "Packed Powder", effect: { damage: 1 }, prerequisiteNodeIds: [] },
    { id: "cannon_small_2", size: "small", branch: "blast", name: "Wider Fuse", effect: { splashRadius: 0.15 }, prerequisiteNodeIds: ["cannon_small_1"] },
    { id: "cannon_medium_1", size: "medium", branch: "blast", name: "Blast Chamber", effect: { splashRadius: 0.35, splashMultiplier: 0.1 }, prerequisiteNodeIds: ["cannon_small_2"] },
    { id: "cannon_small_3", size: "small", branch: "siege", name: "Heavy Barrel", effect: { damage: 2 }, prerequisiteNodeIds: [] },
    { id: "cannon_small_4", size: "small", branch: "siege", name: "Solid Shot", effect: { damage: 2 }, prerequisiteNodeIds: ["cannon_small_3"] },
    { id: "cannon_medium_2", size: "medium", branch: "siege", name: "Impact Frame", effect: { damage: 5 }, prerequisiteNodeIds: ["cannon_small_4"] },
    { id: "cannon_big_1", size: "big", specialist: true, branch: "blast", name: "Demolition Specialist", effect: { splashRadius: 0.6, splashMultiplier: 0.35 }, prerequisiteNodeIds: ["cannon_medium_1"] },
    { id: "cannon_big_2", size: "big", specialist: true, branch: "siege", name: "Siege Specialist", effect: { damage: 10 }, prerequisiteNodeIds: ["cannon_medium_2"] },
  ],
  frost: [
    { id: "frost_small_1", size: "small", branch: "control", name: "Cold Core", effect: { slowDurationMs: 250 }, prerequisiteNodeIds: [] },
    { id: "frost_small_2", size: "small", branch: "control", name: "Biting Wind", effect: { slowAmountDelta: -0.05 }, prerequisiteNodeIds: ["frost_small_1"] },
    { id: "frost_medium_1", size: "medium", branch: "control", name: "Frozen Wake", effect: { slowDurationMs: 600 }, prerequisiteNodeIds: ["frost_small_2"] },
    { id: "frost_small_3", size: "small", branch: "shatter", name: "Ice Edge", effect: { damage: 1 }, prerequisiteNodeIds: [] },
    { id: "frost_small_4", size: "small", branch: "shatter", name: "Crackling Chill", effect: { damage: 1 }, prerequisiteNodeIds: ["frost_small_3"] },
    { id: "frost_medium_2", size: "medium", branch: "shatter", name: "Crystal Fracture", effect: { damage: 4 }, prerequisiteNodeIds: ["frost_small_4"] },
    { id: "frost_big_1", size: "big", specialist: true, branch: "control", name: "Absolute Zero Specialist", effect: { slowDurationMs: 1000, slowAmountDelta: -0.12 }, prerequisiteNodeIds: ["frost_medium_1"] },
    { id: "frost_big_2", size: "big", specialist: true, branch: "shatter", name: "Shatter Specialist", effect: { damage: 7, attackSpeedMs: 40 }, prerequisiteNodeIds: ["frost_medium_2"] },
  ],
};

function getTowerTierUnlockLevel(node) {
  return TOWER_TIER_UNLOCK_LEVELS[node?.size] || 1;
}

function getTowerMasteryNodes(towerType) {
  return TOWER_MASTERY_DEFINITIONS[towerType] || [];
}

function getTowerLoadout(profile, towerType, loadoutId) {
  const loadouts = profile?.towerMasteryLoadouts?.[towerType] || [];
  return loadouts.find((loadout) => String(loadout.id) === String(loadoutId)) || null;
}

function getTowerNodeState(node, selectedIds, purchasedIds, towerLevel) {
  const isSelected = selectedIds.has(node.id);
  const isPurchased = purchasedIds.has(node.id);
  const prerequisitesMet = (node.prerequisiteNodeIds || []).every((reqId) => selectedIds.has(reqId));
  const xpLocked = towerLevel < getTowerTierUnlockLevel(node);
  const active = isSelected && isPurchased && prerequisitesMet && !xpLocked;

  return {
    isSelected,
    isPurchased,
    prerequisitesMet,
    xpLocked,
    active,
  };
}

function getActiveMasteryNodeIdsForTower(tower, profile) {
  const loadout = getTowerLoadout(profile, tower.type, tower.masteryLoadoutId);
  if (!loadout) return [];

  const selectedIds = new Set(loadout.selectedNodeIds || []);
  const purchasedIds = new Set(loadout.purchasedNodeIds || []);
  const nodes = getTowerMasteryNodes(tower.type);
  const towerLevel = getTowerLevelFromXp(tower.xp || 0);

  return nodes
    .filter((node) => getTowerNodeState(node, selectedIds, purchasedIds, towerLevel).active)
    .map((node) => node.id);
}

function getMasteryBonusesForTower(tower, profile) {
  const activeIds = new Set(getActiveMasteryNodeIdsForTower(tower, profile));
  const nodes = getTowerMasteryNodes(tower.type);

  return nodes.reduce(
    (acc, node) => {
      if (!activeIds.has(node.id)) return acc;
      const effect = node.effect || {};
      acc.damage += effect.damage || 0;
      acc.range += effect.range || 0;
      acc.attackSpeedMs += effect.attackSpeedMs || 0;
      acc.splashRadius += effect.splashRadius || 0;
      acc.splashMultiplier += effect.splashMultiplier || 0;
      acc.slowDurationMs += effect.slowDurationMs || 0;
      acc.slowAmountDelta += effect.slowAmountDelta || 0;
      return acc;
    },
    { damage: 0, range: 0, attackSpeedMs: 0, splashRadius: 0, splashMultiplier: 0, slowDurationMs: 0, slowAmountDelta: 0 }
  );
}

function getDefaultPermanentTowerUpgrades() {
  return {
    archer: { damage: 0, range: 0, attackSpeed: 0 },
    cannon: { damage: 0, range: 0, attackSpeed: 0 },
    frost: { damage: 0, range: 0, attackSpeed: 0 },
  };
}

function applyPermanentTowerBonuses(baseTower, permanentBonuses) {
  const bonuses = permanentBonuses || { damage: 0, range: 0, attackSpeed: 0 };
  return {
    ...baseTower,
    damage: baseTower.damage + (bonuses.damage || 0),
    range: baseTower.range + ((bonuses.range || 0) * 0.25),
    cooldownMs: Math.max(150, baseTower.cooldownMs - ((bonuses.attackSpeed || 0) * 20)),
  };
}

function getTowerLevelFromXp(xp) {
  let level = 1;
  let remainingXp = Math.max(0, xp || 0);
  while (remainingXp >= Math.max(25, level * 25)) {
    remainingXp -= Math.max(25, level * 25);
    level += 1;
  }
  return level;
}

function getModifierBonuses(modifier) {
  if (!modifier) {
    return {
      damage: 0,
      range: 0,
      cooldownReductionMs: 0,
      cooldownPenaltyMs: 0,
      splashRadius: 0,
      splashMultiplier: 0,
      bonusGoldOnKill: 0,
      echoChance: 0,
    };
  }

  switch (modifier.type) {
    case "rusted_core":
      return {
        damage: modifier.rolls?.damage || 0,
        range: 0,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: modifier.rolls?.attackSpeedPenaltyMs || 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: 0,
        echoChance: 0,
      };
    case "basic_scope":
      return {
        damage: 0,
        range: (modifier.roll || 0) * 0.18,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: 0,
        echoChance: 0,
      };
    case "cheap_capacitor":
      return {
        damage: 0,
        range: 0,
        cooldownReductionMs: modifier.roll || 0,
        cooldownPenaltyMs: 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: 0,
        echoChance: 0,
      };
    case "midas_touch":
      return {
        damage: 0,
        range: 0,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: modifier.roll || 0,
        echoChance: 0,
      };
    case "volcanic_core":
      return {
        damage: 0,
        range: 0,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: 0,
        splashRadius: (modifier.roll || 0) * 0.12,
        splashMultiplier: (modifier.roll || 0) * 0.05,
        bonusGoldOnKill: 0,
        echoChance: 0,
      };
    case "echo_chamber":
      return {
        damage: 0,
        range: 0,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: 0,
        echoChance: modifier.roll || 0,
      };
    default:
      return {
        damage: 0,
        range: 0,
        cooldownReductionMs: 0,
        cooldownPenaltyMs: 0,
        splashRadius: 0,
        splashMultiplier: 0,
        bonusGoldOnKill: 0,
        echoChance: 0,
      };
  }
}


function getTowerUpgradeCost(tower, stat) {
  const counts = tower?.upgradeCounts || { damage: 0, range: 0, attackSpeed: 0 };

  switch (stat) {
    case "damage":
      return 30 + counts.damage * 15;
    case "range":
      return 55 + counts.range * 30;
    case "attackSpeed":
      return 28 + counts.attackSpeed * 10;
    default:
      return 999999;
  }
}

function getTowerRefundValue(tower) {
  return Math.floor((tower?.totalGoldSpent || tower?.baseCost || 0) * 0.1);
}

function buildGameConfigPayload() {
  return {
    maps: SAMPLE_MAPS,
    towers: TOWER_TYPES,
    bosses: BOSS_TYPES,
    modifiers: MODIFIER_DEFINITIONS,
    towerMasteries: TOWER_MASTERY_DEFINITIONS,
    constants: {
      defaultBossInterval: DEFAULT_BOSS_INTERVAL,
      mapCraftFragmentCost: MAP_CRAFT_FRAGMENT_COST,
      modifierCraftEchoShardCost: MODIFIER_CRAFT_ECHO_SHARD_COST,
      modifierRerollEchoShardCost: MODIFIER_REROLL_ECHO_SHARD_COST,
      modifierEnhanceEchoShardCost: MODIFIER_ENHANCE_ECHO_SHARD_COST,
      towerTierUnlockLevels: TOWER_TIER_UNLOCK_LEVELS,
    },
  };
}

async function handleGetGameConfig(request, env) {
  return json({
    ok: true,
    config: buildGameConfigPayload(),
    serverTime: Date.now(),
  });
}

const MODIFIER_CRAFT_ECHO_SHARD_COST = 25;
const MODIFIER_REROLL_ECHO_SHARD_COST = 10;
const MODIFIER_ENHANCE_ECHO_SHARD_COST = 15;

function rollEchoShardsForModifier(modifier) {
  switch (modifier?.rarity) {
    case "Rare":
      return 4 + Math.floor(Math.random() * 5);
    case "Epic":
      return 10 + Math.floor(Math.random() * 11);
    case "Unique":
      return 25 + Math.floor(Math.random() * 26);
    case "Common":
    default:
      return 1 + Math.floor(Math.random() * 3);
  }
}

function getModifierTierFromWave(wave) {
  return Math.max(1, Math.floor((wave || 1) / 20));
}

function rollModifierForWave(wave, forcedModifierType = null) {
  const tier = getModifierTierFromWave(wave);
  const def = forcedModifierType
    ? MODIFIER_DEFINITIONS.find((item) => item.type === forcedModifierType)
    : MODIFIER_DEFINITIONS[Math.floor(Math.random() * MODIFIER_DEFINITIONS.length)];

  if (!def) return null;

  if (def.rollsConfig) {
    const rolls = {};
    const ranges = {};

    Object.entries(def.rollsConfig).forEach(([key, config]) => {
      const min = config.baseMin + (tier - 1) * config.scaleMin;
      const max = config.baseMax + (tier - 1) * config.scaleMax;
      const roll = Math.floor(Math.random() * (max - min + 1)) + min;
      rolls[key] = roll;
      ranges[key] = { min, max };
    });

    return {
      id: `${def.id}_${Date.now()}_${Math.floor(Math.random() * 100000)}`,
      name: def.name,
      rarity: def.rarity,
      type: def.type,
      rolls,
      ranges,
      enhancementLevel: 0,
    };
  }

  const min = def.baseMin + (tier - 1) * def.scaleMin;
  const max = def.baseMax + (tier - 1) * def.scaleMax;
  const roll = Math.floor(Math.random() * (max - min + 1)) + min;

  return {
    id: `${def.id}_${Date.now()}_${Math.floor(Math.random() * 100000)}`,
    name: def.name,
    rarity: def.rarity,
    type: def.type,
    roll,
    minRoll: min,
    maxRoll: max,
    descriptionTemplate: def.descriptionTemplate,
    enhancementLevel: 0,
  };
}

function rerollModifierValues(modifier, wave) {
  if (!modifier?.type) return modifier;

  const getModifierValueSignature = (item) => {
    if (!item) return "";
    if (item.rolls) return JSON.stringify(item.rolls);
    return JSON.stringify({
      roll: item.roll,
      minRoll: item.minRoll,
      maxRoll: item.maxRoll,
    });
  };

  const currentSignature = getModifierValueSignature(modifier);
  let rerolled = null;

  for (let attempt = 0; attempt < 25; attempt += 1) {
    const candidate = rollModifierForWave(wave, modifier.type);
    if (!candidate) continue;

    const candidateSignature = getModifierValueSignature(candidate);
    rerolled = candidate;

    if (candidateSignature !== currentSignature) {
      break;
    }
  }

  if (!rerolled) return modifier;

  return {
    ...modifier,
    roll: rerolled.roll,
    minRoll: rerolled.minRoll,
    maxRoll: rerolled.maxRoll,
    rolls: rerolled.rolls,
    ranges: rerolled.ranges,
    descriptionTemplate: rerolled.descriptionTemplate || modifier.descriptionTemplate,
    enhancementLevel: modifier.enhancementLevel || 0,
  };
}

function enhanceModifier(modifier) {
  return {
    ...modifier,
    enhancementLevel: (modifier?.enhancementLevel || 0) + 1,
  };
}

function getModifierEnhanceCost(modifier) {
  const level = modifier?.enhancementLevel || 0;
  return MODIFIER_ENHANCE_ECHO_SHARD_COST + level * 10;
}

function getModifierEnhancementMultiplier(modifier) {
  return 1 + ((modifier?.enhancementLevel || 0) * 0.1);
}

function getEnhancedModifierValue(modifier, value) {
  const numericValue = Number(value || 0);
  return Math.round(numericValue * getModifierEnhancementMultiplier(modifier));
}

function getEnhancedModifier(modifier) {
  if (!modifier) return modifier;

  if (modifier.rolls) {
    const nextRolls = {};
    Object.entries(modifier.rolls || {}).forEach(([key, value]) => {
      nextRolls[key] = getEnhancedModifierValue(modifier, value);
    });

    return {
      ...modifier,
      rolls: nextRolls,
    };
  }

  return {
    ...modifier,
    roll: getEnhancedModifierValue(modifier, modifier.roll || 0),
  };
}

function getEnhancedModifierRangeValue(modifier, value) {
  const numericValue = Number(value || 0);
  return Math.round(numericValue * getModifierEnhancementMultiplier(modifier));
}

function getEnhancedModifierRange(modifier, minValue, maxValue) {
  return {
    min: getEnhancedModifierRangeValue(modifier, minValue),
    max: getEnhancedModifierRangeValue(modifier, maxValue),
  };
}

function buildModifierDescriptionParts(modifier) {
  if (!modifier) return [];

  const enhancedModifier = getEnhancedModifier(modifier);

  if (modifier.type === "rusted_core") {
    const damageRange = getEnhancedModifierRange(
      modifier,
      modifier.ranges?.damage?.min ?? 0,
      modifier.ranges?.damage?.max ?? 0
    );
    const speedPenaltyRange = getEnhancedModifierRange(
      modifier,
      modifier.ranges?.attackSpeedPenaltyMs?.min ?? 0,
      modifier.ranges?.attackSpeedPenaltyMs?.max ?? 0
    );

    return [
      { kind: "text", text: "This tower gains " },
      {
        kind: "value",
        value: String(enhancedModifier.rolls?.damage ?? 0),
        range: `(${damageRange.min}-${damageRange.max})`,
      },
      { kind: "text", text: " bonus damage, but attacks " },
      {
        kind: "value",
        value: `${enhancedModifier.rolls?.attackSpeedPenaltyMs ?? 0}ms`,
        range: `(${speedPenaltyRange.min}-${speedPenaltyRange.max})`,
      },
      { kind: "text", text: " more slowly." },
    ];
  }

  if (modifier.type === "basic_scope") {
    const rangeValues = getEnhancedModifierRange(modifier, modifier.minRoll ?? 0, modifier.maxRoll ?? 0);
    return [
      { kind: "text", text: "This tower gains " },
      {
        kind: "value",
        value: String(enhancedModifier.roll ?? 0),
        range: `(${rangeValues.min}-${rangeValues.max})`,
      },
      { kind: "text", text: " bonus range." },
    ];
  }

  if (modifier.type === "cheap_capacitor") {
    const rangeValues = getEnhancedModifierRange(modifier, modifier.minRoll ?? 0, modifier.maxRoll ?? 0);
    return [
      { kind: "text", text: "This tower reduces its cooldown by " },
      {
        kind: "value",
        value: `${enhancedModifier.roll ?? 0}ms`,
        range: `(${rangeValues.min}-${rangeValues.max})`,
      },
      { kind: "text", text: "." },
    ];
  }

  if (modifier.type === "midas_touch") {
    const rangeValues = getEnhancedModifierRange(modifier, modifier.minRoll ?? 0, modifier.maxRoll ?? 0);
    return [
      { kind: "text", text: "Any enemy attacked by this tower drops " },
      {
        kind: "value",
        value: String(enhancedModifier.roll ?? 0),
        range: `(${rangeValues.min}-${rangeValues.max})`,
      },
      { kind: "text", text: " additional gold upon death." },
    ];
  }

  if (modifier.type === "volcanic_core") {
    const rangeValues = getEnhancedModifierRange(modifier, modifier.minRoll ?? 0, modifier.maxRoll ?? 0);
    return [
      { kind: "text", text: "This tower gains " },
      {
        kind: "value",
        value: String(enhancedModifier.roll ?? 0),
        range: `(${rangeValues.min}-${rangeValues.max})`,
      },
      { kind: "text", text: " bonus splash power." },
    ];
  }

  if (modifier.type === "echo_chamber") {
    const rangeValues = getEnhancedModifierRange(modifier, modifier.minRoll ?? 0, modifier.maxRoll ?? 0);
    return [
      { kind: "text", text: "This tower gains a " },
      {
        kind: "value",
        value: `${enhancedModifier.roll ?? 0}%`,
        range: `(${rangeValues.min}-${rangeValues.max})`,
      },
      { kind: "text", text: " chance to repeat its attack." },
    ];
  }

  return [];
}

function buildModifierPreview(modifier) {
  return {
    modifier,
    descriptionParts: buildModifierDescriptionParts(modifier),
    bonuses: getModifierBonuses(getEnhancedModifier(modifier)),
    enhanceCost: getModifierEnhanceCost(modifier),
  };
}

async function handleGetModifierPreview(request, env) {
  const url = new URL(request.url);
  const encoded = url.searchParams.get("modifier");

  if (!encoded) {
    return json({ ok: false, error: "modifier is required." }, { status: 400 });
  }

  let modifier;
  try {
    modifier = JSON.parse(encoded);
  } catch {
    return json({ ok: false, error: "modifier must be valid JSON." }, { status: 400 });
  }

  return json({
    ok: true,
    preview: buildModifierPreview(modifier),
    serverTime: Date.now(),
  });
}

async function loadProfileForModifierAction(env, accountId) {
  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return { error: json({ ok: false, error: "Profile not found." }, { status: 404 }) };
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return { error: json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 }) };
  }

  return { profile };
}

async function saveProfileForModifierAction(env, accountId, profile, now = Date.now()) {
  await env.DB
    .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
    .bind(JSON.stringify(profile), now, accountId)
    .run();

  return now;
}

async function handleCraftModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const loaded = await loadProfileForModifierAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const echoShards = Number(profile?.inventory?.echoShards || 0);
  if (echoShards < MODIFIER_CRAFT_ECHO_SHARD_COST) {
    return json({ ok: false, error: "Not enough Echo Shards." }, { status: 400 });
  }

  const wave = Math.max(1, Number(profile?.stats?.highestWaveReached || 1));
  const modifier = rollModifierForWave(wave);
  if (!modifier) {
    return json({ ok: false, error: "Failed to craft modifier." }, { status: 500 });
  }

  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
      echoShards: echoShards - MODIFIER_CRAFT_ECHO_SHARD_COST,
    },
    modifiers: [modifier, ...(Array.isArray(profile.modifiers) ? profile.modifiers : [])],
  };

  const updatedAt = await saveProfileForModifierAction(env, accountId, nextProfile);

  return json({
    ok: true,
    modifier,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleRerollModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const modifierId = String(body?.modifierId || "");

  if (!accountId || !modifierId) {
    return json({ ok: false, error: "accountId and modifierId are required." }, { status: 400 });
  }

  const loaded = await loadProfileForModifierAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const echoShards = Number(profile?.inventory?.echoShards || 0);
  if (echoShards < MODIFIER_REROLL_ECHO_SHARD_COST) {
    return json({ ok: false, error: "Not enough Echo Shards." }, { status: 400 });
  }

  const modifiers = Array.isArray(profile.modifiers) ? profile.modifiers : [];
  const target = modifiers.find((item) => String(item.id) === modifierId);
  if (!target) {
    return json({ ok: false, error: "Modifier not found." }, { status: 404 });
  }

  const wave = Math.max(1, Number(profile?.stats?.highestWaveReached || 1));
  const rerolled = rerollModifierValues(target, wave);

  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
      echoShards: echoShards - MODIFIER_REROLL_ECHO_SHARD_COST,
    },
    modifiers: modifiers.map((item) =>
      String(item.id) === modifierId ? rerolled : item
    ),
  };

  const updatedAt = await saveProfileForModifierAction(env, accountId, nextProfile);

  return json({
    ok: true,
    modifier: rerolled,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleEnhanceModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const modifierId = String(body?.modifierId || "");

  if (!accountId || !modifierId) {
    return json({ ok: false, error: "accountId and modifierId are required." }, { status: 400 });
  }

  const loaded = await loadProfileForModifierAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const modifiers = Array.isArray(profile.modifiers) ? profile.modifiers : [];
  const target = modifiers.find((item) => String(item.id) === modifierId);
  if (!target) {
    return json({ ok: false, error: "Modifier not found." }, { status: 404 });
  }

  const enhanceCost = getModifierEnhanceCost(target);
  const echoShards = Number(profile?.inventory?.echoShards || 0);
  if (echoShards < enhanceCost) {
    return json({ ok: false, error: "Not enough Echo Shards." }, { status: 400 });
  }

  const enhanced = enhanceModifier(target);
  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
      echoShards: echoShards - enhanceCost,
    },
    modifiers: modifiers.map((item) =>
      String(item.id) === modifierId ? enhanced : item
    ),
  };

  const updatedAt = await saveProfileForModifierAction(env, accountId, nextProfile);

  return json({
    ok: true,
    modifier: enhanced,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleSalvageModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const modifierId = String(body?.modifierId || "");

  if (!accountId || !modifierId) {
    return json({ ok: false, error: "accountId and modifierId are required." }, { status: 400 });
  }

  const loaded = await loadProfileForModifierAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const modifiers = Array.isArray(profile.modifiers) ? profile.modifiers : [];
  const target = modifiers.find((item) => String(item.id) === modifierId);
  if (!target) {
    return json({ ok: false, error: "Modifier not found." }, { status: 404 });
  }

  const shards = rollEchoShardsForModifier(target);
  const nextProfile = {
    ...profile,
    inventory: {
      ...(profile.inventory || {}),
      echoShards: Number(profile?.inventory?.echoShards || 0) + shards,
    },
    modifiers: modifiers.filter((item) => String(item.id) !== modifierId),
  };

  const updatedAt = await saveProfileForModifierAction(env, accountId, nextProfile);

  return json({
    ok: true,
    salvagedModifierId: modifierId,
    echoShardsGained: shards,
    profile: nextProfile,
    updatedAt,
  });
}

function getPermanentUpgradeMaterialForStat(stat) {
  switch (stat) {
    case "damage":
      return "Sunpetal";
    case "attackSpeed":
      return "Ember Shard";
    case "range":
      return "Tideglass";
    default:
      return "Sunpetal";
  }
}

function getPermanentTowerUpgradeCost(towerType, stat, level) {
  switch (stat) {
    case "damage":
      return { gold: 100 + level * 60, materialAmount: 5 + level * 2 };
    case "range":
      return { gold: 180 + level * 90, materialAmount: 8 + level * 3 };
    case "attackSpeed":
      return { gold: 120 + level * 50, materialAmount: 6 + level * 2 };
    default:
      return { gold: 999999, materialAmount: 999999 };
  }
}

async function handleAssignTowerLoadout(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);
  const loadoutId = body?.loadoutId == null ? null : String(body.loadoutId);

  if (!accountId || !runtimeKey || !Number.isFinite(towerId)) {
    return json({ ok: false, error: "accountId, runtimeKey, and towerId are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { profile, activeRun, snapshot } = loaded;
  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const towerIndex = towers.findIndex((tower) => Number(tower.id) === towerId);

  if (towerIndex < 0) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const tower = towers[towerIndex];
  const towerType = tower?.type;
  const loadouts = profile?.towerMasteryLoadouts?.[towerType] || [];

  if (loadoutId !== null) {
    const exists = loadouts.some((loadout) => String(loadout.id) === loadoutId);
    if (!exists) {
      return json({ ok: false, error: "Loadout not found for this tower type." }, { status: 404 });
    }
  }

  const updatedTower = recalculateTowerDerivedStats({
    ...tower,
    masteryLoadoutId: loadoutId,
  });

  const nextTowers = [...towers];
  nextTowers[towerIndex] = updatedTower;

  const now = Date.now();
  const towerLabel = TOWER_TYPES[updatedTower.type]?.label || updatedTower.type;
  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${towerLabel} loadout updated.`,
    lastSimulatedAt: now,
  };

  await env.DB.batch([
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: updatedTower,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function loadProfileForTowerMasteryAction(env, accountId) {
  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return { error: json({ ok: false, error: "Profile not found." }, { status: 404 }) };
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return { error: json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 }) };
  }

  return { profile };
}

async function saveProfileForTowerMasteryAction(env, accountId, profile, now = Date.now()) {
  await env.DB
    .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
    .bind(JSON.stringify(profile), now, accountId)
    .run();

  return now;
}

async function handleSaveTowerLoadout(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const draft = body?.draft;
  const loadoutName = String(body?.loadoutName || "").trim();
  const loadoutId = body?.loadoutId == null ? null : String(body.loadoutId);

  if (!accountId || !towerType || !draft || typeof draft !== "object") {
    return json({ ok: false, error: "accountId, towerType, and draft are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const loaded = await loadProfileForTowerMasteryAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const existingLoadouts = Array.isArray(profile?.towerMasteryLoadouts?.[towerType])
    ? profile.towerMasteryLoadouts[towerType]
    : [];

  const normalizedDraft = {
    id: loadoutId || `loadout_${Date.now()}_${Math.floor(Math.random() * 100000)}`,
    name: loadoutName || `${TOWER_TYPES[towerType]?.label || towerType} Loadout`,
    selectedNodeIds: Array.isArray(draft?.selectedNodeIds) ? draft.selectedNodeIds : [],
    purchasedNodeIds: Array.isArray(draft?.purchasedNodeIds) ? draft.purchasedNodeIds : [],
  };

  const existingIndex = existingLoadouts.findIndex((item) => String(item.id) === String(normalizedDraft.id));
  const nextLoadouts = [...existingLoadouts];

  if (existingIndex >= 0) {
    nextLoadouts[existingIndex] = {
      ...nextLoadouts[existingIndex],
      ...normalizedDraft,
    };
  } else {
    nextLoadouts.unshift(normalizedDraft);
  }

  const nextProfile = {
    ...profile,
    towerMasteryLoadouts: {
      ...(profile.towerMasteryLoadouts || {}),
      [towerType]: nextLoadouts,
    },
  };

  const updatedAt = await saveProfileForTowerMasteryAction(env, accountId, nextProfile);

  return json({
    ok: true,
    towerType,
    loadout: normalizedDraft,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleRenameTowerLoadout(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const loadoutId = String(body?.loadoutId || "");
  const name = String(body?.name || "").trim();

  if (!accountId || !towerType || !loadoutId || !name) {
    return json({ ok: false, error: "accountId, towerType, loadoutId, and name are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const loaded = await loadProfileForTowerMasteryAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const existingLoadouts = Array.isArray(profile?.towerMasteryLoadouts?.[towerType])
    ? profile.towerMasteryLoadouts[towerType]
    : [];

  const index = existingLoadouts.findIndex((item) => String(item.id) === loadoutId);
  if (index < 0) {
    return json({ ok: false, error: "Loadout not found." }, { status: 404 });
  }

  const nextLoadouts = [...existingLoadouts];
  nextLoadouts[index] = {
    ...nextLoadouts[index],
    name,
  };

  const nextProfile = {
    ...profile,
    towerMasteryLoadouts: {
      ...(profile.towerMasteryLoadouts || {}),
      [towerType]: nextLoadouts,
    },
  };

  const updatedAt = await saveProfileForTowerMasteryAction(env, accountId, nextProfile);

  return json({
    ok: true,
    towerType,
    loadout: nextLoadouts[index],
    profile: nextProfile,
    updatedAt,
  });
}

async function handleDeleteTowerLoadout(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const loadoutId = String(body?.loadoutId || "");

  if (!accountId || !towerType || !loadoutId) {
    return json({ ok: false, error: "accountId, towerType, and loadoutId are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const loaded = await loadProfileForTowerMasteryAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const existingLoadouts = Array.isArray(profile?.towerMasteryLoadouts?.[towerType])
    ? profile.towerMasteryLoadouts[towerType]
    : [];

  const nextLoadouts = existingLoadouts.filter((item) => String(item.id) !== loadoutId);
  if (nextLoadouts.length === existingLoadouts.length) {
    return json({ ok: false, error: "Loadout not found." }, { status: 404 });
  }

  const nextProfile = {
    ...profile,
    towerMasteryLoadouts: {
      ...(profile.towerMasteryLoadouts || {}),
      [towerType]: nextLoadouts,
    },
  };

  const updatedAt = await saveProfileForTowerMasteryAction(env, accountId, nextProfile);

  return json({
    ok: true,
    towerType,
    deletedLoadoutId: loadoutId,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleToggleTowerDraftNode(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const nodeId = String(body?.nodeId || "");

  if (!accountId || !towerType || !nodeId) {
    return json({ ok: false, error: "accountId, towerType, and nodeId are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const loaded = await loadProfileForTowerMasteryAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const nodes = getTowerMasteryNodes(towerType);
  const node = nodes.find((item) => item.id === nodeId);
  if (!node) {
    return json({ ok: false, error: "Node not found." }, { status: 404 });
  }

  const towerXp = Number(profile?.towerLevels?.[towerType]?.xp || 0);
  const towerLevel = getTowerLevelFromXp(towerXp);
  const unlockLevel = getTowerTierUnlockLevel(node);

  const currentDraft = profile?.towerMasteryDrafts?.[towerType] || { selectedNodeIds: [], purchasedNodeIds: [] };
  const selectedIds = new Set(currentDraft.selectedNodeIds || []);
  const purchasedIds = new Set(currentDraft.purchasedNodeIds || []);

  const alreadySelected = selectedIds.has(node.id);
  const alreadyPurchased = purchasedIds.has(node.id);
  const prerequisitesMet = (node.prerequisiteNodeIds || []).every((reqId) => selectedIds.has(reqId));

  if (!alreadySelected && towerLevel < unlockLevel) {
    return json({ ok: false, error: `Node unlocks at tower level ${unlockLevel}.` }, { status: 400 });
  }

  if (!alreadySelected && !prerequisitesMet) {
    return json({ ok: false, error: "Prerequisites not met." }, { status: 400 });
  }

  if (!alreadySelected && node.specialist) {
    const specialistIds = nodes.filter((item) => item.specialist).map((item) => item.id);
    const specialistAlreadySelected = (currentDraft.selectedNodeIds || []).some((id) => specialistIds.includes(id));
    if (specialistAlreadySelected) {
      return json({ ok: false, error: "Only one specialist node can be selected." }, { status: 400 });
    }
  }

  let nextGold = Number(profile?.gold || 0);
  if (!alreadySelected && !alreadyPurchased) {
    const goldCost = Number(node.goldCost || 0);
    if (nextGold < goldCost) {
      return json({ ok: false, error: "Not enough gold." }, { status: 400 });
    }
    nextGold -= goldCost;
    purchasedIds.add(node.id);
  }

  if (alreadySelected) {
    selectedIds.delete(node.id);
  } else {
    selectedIds.add(node.id);
  }

  const nextProfile = {
    ...profile,
    gold: nextGold,
    towerMasteryDrafts: {
      ...(profile.towerMasteryDrafts || {}),
      [towerType]: {
        selectedNodeIds: Array.from(selectedIds),
        purchasedNodeIds: Array.from(purchasedIds),
      },
    },
  };

  const updatedAt = await saveProfileForTowerMasteryAction(env, accountId, nextProfile);

  return json({
    ok: true,
    towerType,
    draft: nextProfile.towerMasteryDrafts[towerType],
    gold: nextProfile.gold,
    profile: nextProfile,
    updatedAt,
  });
}

async function handleLoadTowerDraft(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const loadoutId = String(body?.loadoutId || "");

  if (!accountId || !towerType || !loadoutId) {
    return json({ ok: false, error: "accountId, towerType, and loadoutId are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const loaded = await loadProfileForTowerMasteryAction(env, accountId);
  if (loaded.error) return loaded.error;

  const profile = loaded.profile;
  const existingLoadouts = Array.isArray(profile?.towerMasteryLoadouts?.[towerType])
    ? profile.towerMasteryLoadouts[towerType]
    : [];

  const loadout = existingLoadouts.find((item) => String(item.id) === loadoutId);
  if (!loadout) {
    return json({ ok: false, error: "Loadout not found." }, { status: 404 });
  }

  const nextProfile = {
    ...profile,
    towerMasteryDrafts: {
      ...(profile.towerMasteryDrafts || {}),
      [towerType]: {
        selectedNodeIds: Array.isArray(loadout.selectedNodeIds) ? loadout.selectedNodeIds : [],
        purchasedNodeIds: Array.isArray(loadout.purchasedNodeIds) ? loadout.purchasedNodeIds : [],
      },
    },
  };

  const updatedAt = await saveProfileForTowerMasteryAction(env, accountId, nextProfile);

  return json({
    ok: true,
    towerType,
    draft: nextProfile.towerMasteryDrafts[towerType],
    profile: nextProfile,
    updatedAt,
  });
}

async function handlePermanentTowerUpgrade(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const towerType = String(body?.towerType || "");
  const stat = String(body?.stat || "");

  if (!accountId || !towerType || !stat) {
    return json({ ok: false, error: "accountId, towerType, and stat are required." }, { status: 400 });
  }

  if (!TOWER_TYPES[towerType]) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  if (!["damage", "range", "attackSpeed"].includes(stat)) {
    return json({ ok: false, error: "Invalid permanent upgrade stat." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const currentLevels = profile?.permanentTowerUpgrades?.[towerType] || { damage: 0, range: 0, attackSpeed: 0 };
  const currentLevel = Number(currentLevels?.[stat] || 0);
  const cost = getPermanentTowerUpgradeCost(towerType, stat, currentLevel);
  const materialName = getPermanentUpgradeMaterialForStat(stat);
  const materialOwned = Number(profile?.inventory?.[materialName] || 0);
  const goldOwned = Number(profile?.gold || 0);

  if (goldOwned < cost.gold) {
    return json({ ok: false, error: "Not enough gold." }, { status: 400 });
  }

  if (materialOwned < cost.materialAmount) {
    return json({ ok: false, error: `Not enough ${materialName}.` }, { status: 400 });
  }

  const nextProfile = {
    ...profile,
    gold: goldOwned - cost.gold,
    inventory: {
      ...(profile.inventory || {}),
      [materialName]: materialOwned - cost.materialAmount,
    },
    permanentTowerUpgrades: {
      ...(profile.permanentTowerUpgrades || {}),
      [towerType]: {
        ...(profile.permanentTowerUpgrades?.[towerType] || { damage: 0, range: 0, attackSpeed: 0 }),
        [stat]: currentLevel + 1,
      },
    },
  };

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  const now = Date.now();
  const statements = [
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
  ];

  for (const row of activeRunsRows?.results || []) {
    const activeRun = {
      slotIndex: row.slot_index,
      runtimeKey: row.runtime_key,
      mapId: row.map_id,
      sourceMapId: row.source_map_id,
      isCrafted: Boolean(row.is_crafted),
      craftedMapId: row.crafted_map_id,
      status: row.status,
      startedAt: row.started_at,
      lastSimulatedAt: row.last_simulated_at,
      selected: Boolean(row.selected),
      opened: Boolean(row.opened),
      wave: row.wave,
      highestWaveReached: row.highest_wave_reached,
      updatedAt: row.updated_at,
    };

    const runtimeMap = getRuntimeMapFromActiveRun(activeRun, nextProfile);
    if (!runtimeMap) continue;

    const snapshotRow = await env.DB
      .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
      .bind(accountId, activeRun.runtimeKey)
      .first();

    let snapshot;
    try {
      snapshot = snapshotRow?.snapshot_json
        ? JSON.parse(snapshotRow.snapshot_json)
        : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, now);
    } catch {
      snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, now);
    }

    const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
    const hasMatchingTowerType = towers.some((tower) => tower?.type === towerType);
    if (!hasMatchingTowerType) continue;

    const updatedTowers = towers.map((tower) => {
      if (tower?.type !== towerType) return tower;

      const updated = { ...tower };
      if (stat === "damage") updated.baseDamage = Number(updated.baseDamage ?? TOWER_TYPES[towerType].damage ?? 0) + 1;
      if (stat === "range") updated.baseRange = Number(updated.baseRange ?? TOWER_TYPES[towerType].range ?? 0) + 0.25;
      if (stat === "attackSpeed") {
        updated.baseCooldownMs = Math.max(
          150,
          Number(updated.baseCooldownMs ?? TOWER_TYPES[towerType].cooldownMs ?? 1000) - 20
        );
      }
      return recalculateTowerDerivedStats(updated);
    });

    const nextSnapshot = {
      ...snapshot,
      towers: updatedTowers,
      status: `${TOWER_TYPES[towerType]?.label || towerType} permanent ${stat} upgraded.`,
      lastSimulatedAt: now,
    };

    statements.push(
      env.DB
        .prepare(
          `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(account_id, run_key) DO UPDATE SET
             snapshot_json = excluded.snapshot_json,
             updated_at = excluded.updated_at`
        )
        .bind(accountId, activeRun.runtimeKey, JSON.stringify(nextSnapshot), now)
    );

    statements.push(
      env.DB
        .prepare(
          `UPDATE active_runs
           SET last_simulated_at = ?, updated_at = ?
           WHERE account_id = ? AND runtime_key = ?`
        )
        .bind(now, now, accountId, activeRun.runtimeKey)
    );
  }

  await env.DB.batch(statements);

  return json({
    ok: true,
    towerType,
    stat,
    profile: nextProfile,
    updatedAt: now,
  });
}

function isUniqueModifier(modifier) {
  return modifier?.rarity === "Unique";
}

function canEquipModifierOnMap(modifier, towers, towerIdToIgnore = null) {
  if (!isUniqueModifier(modifier)) return true;

  const usedTypes = new Set(
    (towers || [])
      .filter((tower) => Number(tower.id) !== Number(towerIdToIgnore))
      .map((tower) => tower?.modifier?.type)
      .filter(Boolean)
  );

  return !usedTypes.has(modifier.type);
}

async function handleEquipModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);
  const modifierId = String(body?.modifierId || "");

  if (!accountId || !runtimeKey || !Number.isFinite(towerId) || !modifierId) {
    return json({ ok: false, error: "accountId, runtimeKey, towerId, and modifierId are required." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const towerIndex = towers.findIndex((tower) => Number(tower.id) === towerId);
  if (towerIndex < 0) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const modifiers = Array.isArray(profile.modifiers) ? profile.modifiers : [];
  const modifierIndex = modifiers.findIndex((modifier) => String(modifier.id) === modifierId);
  if (modifierIndex < 0) {
    return json({ ok: false, error: "Modifier not found in inventory." }, { status: 404 });
  }

  const modifier = modifiers[modifierIndex];
  if (!canEquipModifierOnMap(modifier, towers, towerId)) {
    return json({ ok: false, error: `Only one ${modifier.name} can be equipped on this map.` }, { status: 400 });
  }

  const targetTower = towers[towerIndex];
  const returnedModifier = targetTower?.modifier || null;

  const updatedTower = recalculateTowerDerivedStats({
    ...targetTower,
    modifier,
  }, profile);

  const nextTowers = [...towers];
  nextTowers[towerIndex] = updatedTower;

  const nextModifiers = modifiers.filter((item) => String(item.id) !== modifierId);
  if (returnedModifier) {
    nextModifiers.unshift(returnedModifier);
  }

  const now = Date.now();
  const towerLabel = TOWER_TYPES[updatedTower.type]?.label || updatedTower.type;

  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${modifier.name} equipped on ${towerLabel}.`,
    lastSimulatedAt: now,
  };

  const nextProfile = {
    ...profile,
    modifiers: nextModifiers,
  };

  await env.DB.batch([
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: updatedTower,
    profile: nextProfile,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}


function getEnemyTypeDefinitionsForAdmin() {
  return {
    pink_circle: {
      key: "pink_circle",
      label: "Pink Circle",
      speed: 1 / 6,
      size: 16,
      colorClass: "bg-rose-400 border-rose-100",
      rewardScale: 1,
      hpScale: 1,
    },
    dark_green_circle: {
      key: "dark_green_circle",
      label: "Dark Green Circle",
      speed: 1 / 4,
      size: 12,
      colorClass: "bg-green-800 border-green-300",
      rewardScale: 1.15,
      hpScale: 0.9,
    },
    light_purple_circle: {
      key: "light_purple_circle",
      label: "Light Purple Circle",
      speed: 1 / 8,
      size: 22,
      colorClass: "bg-violet-300 border-violet-100",
      rewardScale: 1.45,
      hpScale: 2.0,
    },
  };
}

const ADMIN_BOSS_TYPES = {
  guardian: {
    key: "guardian",
    label: "Cave Guardian",
    speed: 1 / 10,
    size: 30,
    colorClass: "bg-stone-300 border-stone-100",
    hpMultiplier: 14,
    rewardMultiplier: 10,
  },
  broodmother: {
    key: "broodmother",
    label: "Broodmother",
    speed: 1 / 9,
    size: 28,
    colorClass: "bg-fuchsia-500 border-fuchsia-200",
    hpMultiplier: 12,
    rewardMultiplier: 11,
  },
  titan: {
    key: "titan",
    label: "Stone Titan",
    speed: 1 / 12,
    size: 34,
    colorClass: "bg-slate-400 border-slate-100",
    hpMultiplier: 18,
    rewardMultiplier: 13,
  },
};

function getBossIntervalFromMap(runtimeMap) {
  return Number(runtimeMap?.bossConfig?.interval || 20);
}

function isBossWaveForAdmin(wave, runtimeMap) {
  const interval = getBossIntervalFromMap(runtimeMap);
  return wave > 0 && wave % interval === 0;
}

function getBossTierForAdmin(wave, runtimeMap) {
  const interval = getBossIntervalFromMap(runtimeMap);
  return Math.max(1, Math.floor(wave / interval));
}

function getBossKeyForAdmin(wave, runtimeMap) {
  const rotation = runtimeMap?.bossConfig?.rotation || [];
  if (!rotation.length) return null;
  const tier = getBossTierForAdmin(wave, runtimeMap);
  return rotation[(tier - 1) % rotation.length] || null;
}

function createBossEnemyForAdmin({ wave, runtimeMap, enemyId }) {
  const bossKey = getBossKeyForAdmin(wave, runtimeMap);
  const bossType = ADMIN_BOSS_TYPES[bossKey];
  if (!bossType) return null;

  const hpScale = (1 + (wave - 1) * 0.18) * bossType.hpMultiplier * (1 + (getBossTierForAdmin(wave, runtimeMap) - 1) * 0.35);
  const rewardScale = (1 + (wave - 1) * 0.06) * bossType.rewardMultiplier * (1 + (getBossTierForAdmin(wave, runtimeMap) - 1) * 0.2);
  const baseHp = 24;
  const baseReward = 8;
  const hp = Math.max(1, Math.round(baseHp * hpScale));
  const reward = Math.max(1, Math.round(baseReward * rewardScale));

  return {
    id: enemyId,
    type: bossType.key,
    bossType: bossType.key,
    label: bossType.label,
    hp,
    maxHp: hp,
    reward,
    dropChance: 1,
    pathPosition: 0,
    speed: bossType.speed,
    size: bossType.size,
    colorClass: bossType.colorClass,
    slowedUntil: 0,
    slowMultiplier: 1,
    isBoss: true,
    bossTier: getBossTierForAdmin(wave, runtimeMap),
    waveSpawned: wave,
  };
}

async function loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey) {
  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return { error: json({ ok: false, error: "Profile not found." }, { status: 404 }) };
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return { error: json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 }) };
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return { error: json({ ok: false, error: "Active run not found." }, { status: 404 }) };
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return { error: json({ ok: false, error: "Runtime map not found." }, { status: 404 }) };
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  return { profile, activeRun, runtimeMap, snapshot };
}

async function saveAdminUpdatedSnapshot(env, accountId, runtimeKey, activeRun, nextSnapshot, now) {
  await env.DB.batch([
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET wave = ?, highest_wave_reached = ?, last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(
        Number(nextSnapshot.wave || activeRun.wave || 1),
        Number(nextSnapshot.highestWaveReached || activeRun.highestWaveReached || nextSnapshot.wave || 1),
        now,
        now,
        accountId,
        runtimeKey
      ),
  ]);
}

async function handleAdminSetWave(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const wave = Math.max(1, Number(body?.wave || 1));

  if (!accountId || !runtimeKey || !Number.isFinite(wave)) {
    return json({ ok: false, error: "accountId, runtimeKey, and wave are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { activeRun, snapshot } = loaded;
  const now = Date.now();

  const nextSnapshot = {
    ...snapshot,
    wave,
    highestWaveReached: Math.max(Number(snapshot.highestWaveReached || 1), wave),
    enemies: [],
    popups: [],
    spawnTimerMs: 0,
    enemiesSpawnedInWave: 0,
    status: `Admin set wave to ${wave}. Towers kept in place.`,
    lastSimulatedAt: now,
  };

  await saveAdminUpdatedSnapshot(env, accountId, runtimeKey, activeRun, nextSnapshot, now);

  return json({
    ok: true,
    runtimeKey,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleAdminClearEnemies(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { activeRun, snapshot } = loaded;
  const now = Date.now();

  const nextSnapshot = {
    ...snapshot,
    enemies: [],
    popups: [],
    status: "Cleared all active enemies. Towers kept in place.",
    lastSimulatedAt: now,
  };

  await saveAdminUpdatedSnapshot(env, accountId, runtimeKey, activeRun, nextSnapshot, now);

  return json({
    ok: true,
    runtimeKey,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleAdminSpawnTestEnemy(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { activeRun, runtimeMap, snapshot } = loaded;
  const ENEMY_TYPES = getEnemyTypeDefinitionsForAdmin();
  const spawnType = ENEMY_TYPES.pink_circle;
  const wave = Number(snapshot.wave || activeRun.wave || 1);
  const hpScale = (1 + (wave - 1) * 0.18) * spawnType.hpScale;
  const rewardScale = (1 + (wave - 1) * 0.06) * spawnType.rewardScale;
  const enemyId = Math.max(1, ...((snapshot.enemies || []).map((enemy) => Number(enemy.id) || 0)), 0) + 1;
  const baseHp = 24;
  const baseReward = 8;
  const now = Date.now();

  const enemy = {
    id: enemyId,
    type: spawnType.key,
    label: `${spawnType.label} (Test)`,
    hp: Math.round(baseHp * hpScale),
    maxHp: Math.round(baseHp * hpScale),
    reward: Math.max(1, Math.round(baseReward * rewardScale)),
    dropChance: 0.25,
    pathPosition: 0,
    speed: spawnType.speed,
    size: spawnType.size,
    colorClass: spawnType.colorClass,
    slowedUntil: 0,
    slowMultiplier: 1,
    isBoss: false,
  };

  const nextSnapshot = {
    ...snapshot,
    enemies: [...(Array.isArray(snapshot.enemies) ? snapshot.enemies : []), enemy],
    status: `Spawned test enemy on wave ${wave}.`,
    lastSimulatedAt: now,
  };

  await saveAdminUpdatedSnapshot(env, accountId, runtimeKey, activeRun, nextSnapshot, now);

  return json({
    ok: true,
    runtimeKey,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleAdminSpawnBoss(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { activeRun, runtimeMap, snapshot } = loaded;
  const currentWave = Number(snapshot.wave || activeRun.wave || 1);
  const interval = getBossIntervalFromMap(runtimeMap);
  const bossWave = isBossWaveForAdmin(currentWave, runtimeMap)
    ? currentWave
    : Math.ceil(currentWave / interval) * interval;

  const enemyId = Math.max(1, ...((snapshot.enemies || []).map((enemy) => Number(enemy.id) || 0)), 0) + 1;
  const bossEnemy = createBossEnemyForAdmin({
    wave: bossWave,
    runtimeMap,
    enemyId,
  });

  if (!bossEnemy) {
    return json({ ok: false, error: "No boss could be generated for this map." }, { status: 400 });
  }

  const now = Date.now();
  const nextSnapshot = {
    ...snapshot,
    enemies: [...(Array.isArray(snapshot.enemies) ? snapshot.enemies : []), bossEnemy],
    status: `Spawned boss test: ${bossEnemy.label} (Tier ${bossEnemy.bossTier})`,
    lastSimulatedAt: now,
  };

  await saveAdminUpdatedSnapshot(env, accountId, runtimeKey, activeRun, nextSnapshot, now);

  return json({
    ok: true,
    runtimeKey,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleUnequipModifier(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);

  if (!accountId || !runtimeKey || !Number.isFinite(towerId)) {
    return json({ ok: false, error: "accountId, runtimeKey, and towerId are required." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const towerIndex = towers.findIndex((tower) => Number(tower.id) === towerId);
  if (towerIndex < 0) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const targetTower = towers[towerIndex];
  if (!targetTower?.modifier) {
    return json({ ok: false, error: "Tower has no modifier equipped." }, { status: 400 });
  }

  const returnedModifier = targetTower.modifier;
  const updatedTower = recalculateTowerDerivedStats({
    ...targetTower,
    modifier: null,
  }, profile);

  const nextTowers = [...towers];
  nextTowers[towerIndex] = updatedTower;

  const nextProfile = {
    ...profile,
    modifiers: [returnedModifier, ...(Array.isArray(profile.modifiers) ? profile.modifiers : [])],
  };

  const now = Date.now();
  const towerLabel = TOWER_TYPES[updatedTower.type]?.label || updatedTower.type;

  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${returnedModifier.name} returned to inventory from ${towerLabel}.`,
    lastSimulatedAt: now,
  };

  await env.DB.batch([
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: updatedTower,
    profile: nextProfile,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleSetTowerPriority(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);
  const priority = String(body?.priority || "");

  if (!accountId || !runtimeKey || !Number.isFinite(towerId) || !priority) {
    return json({ ok: false, error: "accountId, runtimeKey, towerId, and priority are required." }, { status: 400 });
  }

  if (!["first", "closest", "last"].includes(priority)) {
    return json({ ok: false, error: "Invalid target priority." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const towerIndex = towers.findIndex((tower) => Number(tower.id) === towerId);
  if (towerIndex < 0) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const updatedTower = recalculateTowerDerivedStats({
    ...towers[towerIndex],
    targetPriority: priority,
  }, profile);

  const nextTowers = [...towers];
  nextTowers[towerIndex] = updatedTower;

  const now = Date.now();
  const towerLabel = TOWER_TYPES[updatedTower.type]?.label || updatedTower.type;

  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${towerLabel} priority set to ${priority}.`,
    lastSimulatedAt: now,
  };

  await env.DB.batch([
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: updatedTower,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleRemoveTower(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);

  if (!accountId || !runtimeKey || !Number.isFinite(towerId)) {
    return json({ ok: false, error: "accountId, runtimeKey, and towerId are required." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const tower = towers.find((item) => Number(item.id) === towerId);
  if (!tower) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const refund = getTowerRefundValue(tower);
  const nextTowers = towers.filter((item) => Number(item.id) !== towerId);

  const now = Date.now();
  const towerLabel = TOWER_TYPES[tower.type]?.label || tower.type;

  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${towerLabel} removed for ${refund} gold refund.`,
    lastSimulatedAt: now,
  };

  const nextProfile = {
    ...profile,
    gold: Number(profile.gold || 0) + refund,
  };

  await env.DB.batch([
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    removedTowerId: towerId,
    refund,
    profile: nextProfile,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

async function handleUpgradeTower(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerId = Number(body?.towerId);
  const stat = String(body?.stat || "");

  if (!accountId || !runtimeKey || !Number.isFinite(towerId) || !stat) {
    return json({ ok: false, error: "accountId, runtimeKey, towerId, and stat are required." }, { status: 400 });
  }

  if (!["damage", "range", "attackSpeed"].includes(stat)) {
    return json({ ok: false, error: "Invalid upgrade stat." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const towerIndex = towers.findIndex((tower) => Number(tower.id) === towerId);
  if (towerIndex < 0) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  const tower = towers[towerIndex];
  const cost = getTowerUpgradeCost(tower, stat);

  if (Number(profile.gold || 0) < cost) {
    return json({ ok: false, error: `Not enough gold for ${stat} upgrade.` }, { status: 400 });
  }

  const nextCounts = {
    ...(tower.upgradeCounts || { damage: 0, range: 0, attackSpeed: 0 }),
  };
  nextCounts[stat] = Number(nextCounts[stat] || 0) + 1;

  const updatedTower = recalculateTowerDerivedStats({
    ...tower,
    upgradeCounts: nextCounts,
    totalGoldSpent: Number(tower.totalGoldSpent || tower.baseCost || 0) + cost,
  }, profile);

  const nextTowers = [...towers];
  nextTowers[towerIndex] = updatedTower;

  const now = Date.now();
  const nextSnapshot = {
    ...snapshot,
    towers: nextTowers,
    status: `${TOWER_TYPES[tower.type]?.label || tower.type} upgraded: ${stat}.`,
    lastSimulatedAt: now,
  };

  const nextProfile = {
    ...profile,
    gold: Number(profile.gold || 0) - cost,
  };

  await env.DB.batch([
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: updatedTower,
    profile: nextProfile,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

function recalculateTowerDerivedStats(tower, profile = null) {
  const counts = tower.upgradeCounts || { damage: 0, range: 0, attackSpeed: 0 };
  const masteryBonuses = profile ? getMasteryBonusesForTower(tower, profile) : { damage: 0, range: 0, attackSpeedMs: 0, splashRadius: 0, splashMultiplier: 0, slowDurationMs: 0, slowAmountDelta: 0 };
  const activeMasteryNodeIds = profile ? getActiveMasteryNodeIdsForTower(tower, profile) : (Array.isArray(tower.activeMasteryNodeIds) ? tower.activeMasteryNodeIds : []);
  const modifierBonuses = getModifierBonuses(tower.modifier);

  const typeStats = TOWER_TYPES[tower.type] || {};
  const baseDamage = tower.baseDamage ?? typeStats.damage ?? 0;
  const baseRange = tower.baseRange ?? typeStats.range ?? 0;
  const baseCooldownMs = tower.baseCooldownMs ?? typeStats.cooldownMs ?? 1000;
  const baseSplashRadius = tower.baseSplashRadius ?? (tower.type === "cannon" ? 1.0 : 0);
  const baseSplashMultiplier = tower.baseSplashMultiplier ?? (tower.type === "cannon" ? 0.35 : 0);
  const baseSlowAmount = tower.baseSlowAmount ?? typeStats.slowAmount ?? null;
  const baseSlowDurationMs = tower.baseSlowDurationMs ?? typeStats.slowDurationMs ?? null;

  const nextSlowAmount =
    baseSlowAmount == null ? null : Math.max(0.2, baseSlowAmount + masteryBonuses.slowAmountDelta);

  return {
    ...tower,
    baseDamage,
    baseRange,
    baseCooldownMs,
    baseSplashRadius,
    baseSplashMultiplier,
    baseSlowAmount,
    baseSlowDurationMs,
    level: tower.level || getTowerLevelFromXp(tower.xp || 0),
    damage: baseDamage + counts.damage + masteryBonuses.damage + modifierBonuses.damage,
    range: baseRange + counts.range * 0.35 + masteryBonuses.range + modifierBonuses.range,
    cooldownMs: Math.max(
      150,
      baseCooldownMs
        - counts.attackSpeed * 25
        - masteryBonuses.attackSpeedMs
        - modifierBonuses.cooldownReductionMs
        + modifierBonuses.cooldownPenaltyMs
    ),
    splashRadius: Math.max(0, baseSplashRadius + masteryBonuses.splashRadius + modifierBonuses.splashRadius),
    splashMultiplier: Math.max(0, baseSplashMultiplier + masteryBonuses.splashMultiplier + modifierBonuses.splashMultiplier),
    slowAmount: nextSlowAmount,
    slowDurationMs: baseSlowDurationMs == null ? null : baseSlowDurationMs + masteryBonuses.slowDurationMs,
    bonusGoldOnKill: modifierBonuses.bonusGoldOnKill || 0,
    echoChance: modifierBonuses.echoChance || 0,
    activeMasteryNodeIds,
  };
}

async function handlePlaceTower(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const runtimeKey = String(body?.runtimeKey || "");
  const towerType = String(body?.towerType || "");
  const x = Number(body?.x);
  const y = Number(body?.y);

  if (!accountId || !runtimeKey || !towerType || !Number.isInteger(x) || !Number.isInteger(y)) {
    return json({ ok: false, error: "accountId, runtimeKey, towerType, x, and y are required." }, { status: 400 });
  }

  const towerConfig = TOWER_TYPES[towerType];
  if (!towerConfig) {
    return json({ ok: false, error: "Unknown tower type." }, { status: 400 });
  }

  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunRow = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ? AND runtime_key = ?
       LIMIT 1`
    )
    .bind(accountId, runtimeKey)
    .first();

  if (!activeRunRow) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const activeRun = {
    slotIndex: activeRunRow.slot_index,
    runtimeKey: activeRunRow.runtime_key,
    mapId: activeRunRow.map_id,
    sourceMapId: activeRunRow.source_map_id,
    isCrafted: Boolean(activeRunRow.is_crafted),
    craftedMapId: activeRunRow.crafted_map_id,
    status: activeRunRow.status,
    startedAt: activeRunRow.started_at,
    lastSimulatedAt: activeRunRow.last_simulated_at,
    selected: Boolean(activeRunRow.selected),
    opened: Boolean(activeRunRow.opened),
    wave: activeRunRow.wave,
    highestWaveReached: activeRunRow.highest_wave_reached,
    updatedAt: activeRunRow.updated_at,
  };

  const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
  if (!runtimeMap) {
    return json({ ok: false, error: "Runtime map not found." }, { status: 404 });
  }

  const tileType = getRuntimeMapTileType(runtimeMap, x, y);
  if (tileType !== "build") {
    return json({
      ok: false,
      error: `You can only place towers on build tiles. Server saw ${x},${y} as ${tileType || "unknown"} on ${runtimeMap?.id || runtimeMap?.mapId || runtimeKey}.`,
    }, { status: 400 });
  }

  if (towerConfig.unlockLevel && Number(profile.level || 1) < towerConfig.unlockLevel) {
    return json({ ok: false, error: `${towerConfig.label} unlocks at player level ${towerConfig.unlockLevel}.` }, { status: 400 });
  }

  if (Number(profile.gold || 0) < towerConfig.cost) {
    return json({ ok: false, error: `Not enough gold for ${towerConfig.label}.` }, { status: 400 });
  }

  const snapshotRow = await env.DB
    .prepare("SELECT snapshot_json FROM run_snapshots WHERE account_id = ? AND run_key = ?")
    .bind(accountId, runtimeKey)
    .first();

  let snapshot;
  try {
    snapshot = snapshotRow?.snapshot_json
      ? JSON.parse(snapshotRow.snapshot_json)
      : createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  } catch {
    snapshot = createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, Date.now());
  }

  const towers = Array.isArray(snapshot.towers) ? snapshot.towers : [];
  const occupied = towers.some((tower) => Number(tower.x) === x && Number(tower.y) === y);
  if (occupied) {
    return json({ ok: false, error: "There is already a tower on that tile." }, { status: 400 });
  }

  const permanentBonuses =
    profile.permanentTowerUpgrades?.[towerConfig.key] ||
    getDefaultPermanentTowerUpgrades()[towerConfig.key] ||
    { damage: 0, range: 0, attackSpeed: 0 };

  const nextTowerId =
    Math.max(1, ...towers.map((tower) => Number(tower.id) || 0), 0) + 1;

  const placedTower = recalculateTowerDerivedStats(
    applyPermanentTowerBonuses(
      {
        id: nextTowerId,
        type: towerConfig.key,
        x,
        y,
        damage: towerConfig.damage,
        range: towerConfig.range,
        cooldownMs: towerConfig.cooldownMs,
        lastShotAt: 0,
        kills: 0,
        lifetimeGoldEarned: 0,
        targetPriority: "first",
        slowAmount: towerConfig.slowAmount || null,
        slowDurationMs: towerConfig.slowDurationMs || null,
        baseCost: towerConfig.cost,
        totalGoldSpent: towerConfig.cost,
        upgradeCounts: {
          damage: 0,
          range: 0,
          attackSpeed: 0,
        },
        modifier: null,
        masteryLoadoutId: null,
        activeMasteryNodeIds: [],
        xp: 0,
        level: 1,
        baseDamage: towerConfig.damage,
        baseRange: towerConfig.range,
        baseCooldownMs: towerConfig.cooldownMs,
        baseSplashRadius: towerConfig.key === "cannon" ? 1.0 : 0,
        baseSplashMultiplier: towerConfig.key === "cannon" ? 0.35 : 0,
        baseSlowAmount: towerConfig.slowAmount || null,
        baseSlowDurationMs: towerConfig.slowDurationMs || null,
      },
      permanentBonuses
    ),
    profile
  );

  const now = Date.now();
  const nextSnapshot = {
    ...snapshot,
    towers: [...towers, placedTower],
    status: `${towerConfig.label} placed at (${x}, ${y}).`,
    mapId: runtimeKey,
    lastSimulatedAt: now,
  };

  const nextProfile = {
    ...profile,
    gold: Number(profile.gold || 0) - towerConfig.cost,
  };

  await env.DB.batch([
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(account_id, run_key) DO UPDATE SET
           snapshot_json = excluded.snapshot_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, runtimeKey, JSON.stringify(nextSnapshot), now),
    env.DB
      .prepare(
        `UPDATE active_runs
         SET last_simulated_at = ?, updated_at = ?
         WHERE account_id = ? AND runtime_key = ?`
      )
      .bind(now, now, accountId, runtimeKey),
  ]);

  return json({
    ok: true,
    runtimeKey,
    tower: placedTower,
    profile: nextProfile,
    snapshot: nextSnapshot,
    updatedAt: now,
  });
}

function createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, now) {
  return {
    xp: 0,
    wave: Number(activeRun?.wave || 1),
    castleLives: Number(runtimeMap?.maxLeaks || 20),
    enemies: [],
    towers: [],
    popups: [],
    drops: {},
    status: `Resuming ${runtimeMap?.name || activeRun?.runtimeKey || "run"} from backend active run.`,
    isRunning: activeRun?.status === "running",
    mapId: activeRun?.runtimeKey || runtimeMap?.id || activeRun?.mapId || "grasslands",
    startedAt: Number(activeRun?.startedAt || activeRun?.lastSimulatedAt || now),
    lastSimulatedAt: Number(activeRun?.lastSimulatedAt || activeRun?.startedAt || now),
    elapsedMs: 0,
    spawnTimerMs: 0,
    enemiesSpawnedInWave: 0,
    tracker: {
      goldEarned: 0,
      playerXpEarned: 0,
    },
    highestWaveReached: Number(activeRun?.highestWaveReached || activeRun?.wave || 1),
    dropLog: [],
    killLog: [],
    damageLog: [],
    leakLog: [],
    isDamageLogPaused: false,
    isKillLogPaused: false,
    isLeakLogPaused: false,
  };
}

async function handleStartActiveRun(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const slotIndex = body?.slotIndex;

  if (!accountId || slotIndex == null) {
    return json({ ok: false, error: "accountId and slotIndex are required." }, { status: 400 });
  }

  const existing = await getExistingActiveRunForAction(env, accountId, slotIndex);
  if (!existing) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `UPDATE active_runs
       SET status = ?, last_simulated_at = ?, updated_at = ?
       WHERE account_id = ? AND slot_index = ?`
    )
    .bind("running", updatedAt, updatedAt, accountId, Number(slotIndex))
    .run();

  return json({
    ok: true,
    updatedAt,
    activeRun: {
      ...normalizeActiveRunRow(existing, updatedAt),
      status: "running",
      lastSimulatedAt: updatedAt,
      updatedAt,
    },
  });
}

async function handleStopActiveRun(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const slotIndex = body?.slotIndex;

  if (!accountId || slotIndex == null) {
    return json({ ok: false, error: "accountId and slotIndex are required." }, { status: 400 });
  }

  const existing = await getExistingActiveRunForAction(env, accountId, slotIndex);
  if (!existing) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `UPDATE active_runs
       SET status = ?, updated_at = ?
       WHERE account_id = ? AND slot_index = ?`
    )
    .bind("stopped", updatedAt, accountId, Number(slotIndex))
    .run();

  return json({
    ok: true,
    updatedAt,
    activeRun: {
      ...normalizeActiveRunRow(existing, updatedAt),
      status: "stopped",
      updatedAt,
    },
  });
}

async function handleResetActiveRun(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;
  const slotIndex = body?.slotIndex;

  if (!accountId || slotIndex == null) {
    return json({ ok: false, error: "accountId and slotIndex are required." }, { status: 400 });
  }

  const existing = await getExistingActiveRunForAction(env, accountId, slotIndex);
  if (!existing) {
    return json({ ok: false, error: "Active run not found." }, { status: 404 });
  }

  const updatedAt = Date.now();

  await env.DB
    .prepare(
      `UPDATE active_runs
       SET status = ?, started_at = ?, last_simulated_at = ?, wave = ?, highest_wave_reached = ?, updated_at = ?
       WHERE account_id = ? AND slot_index = ?`
    )
    .bind("running", updatedAt, updatedAt, 1, 1, updatedAt, accountId, Number(slotIndex))
    .run();

  return json({
    ok: true,
    updatedAt,
    activeRun: {
      ...normalizeActiveRunRow(existing, updatedAt),
      status: "running",
      startedAt: updatedAt,
      lastSimulatedAt: updatedAt,
      wave: 1,
      highestWaveReached: 1,
      updatedAt,
    },
  });
}



async function handleCatchupActiveRuns(request, env) {
  const body = await readJson(request);
  const accountId = body?.accountId;

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const result = await runBackendCatchupForAccount(env, accountId);
  if (!result?.ok) {
    return json({ ok: false, error: result?.error || "Catchup failed." }, { status: 400 });
  }

  return json(result);
}

async function handleCatchupActiveRuns_legacy_do_not_use(request, env) {
  const profileRow = await env.DB
    .prepare("SELECT profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  const activeRunsRows = await env.DB
    .prepare(
      `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
              status, started_at, last_simulated_at, selected, opened, wave,
              highest_wave_reached, updated_at
       FROM active_runs
       WHERE account_id = ?
       ORDER BY slot_index ASC`
    )
    .bind(accountId)
    .all();

  const snapshotRows = await env.DB
    .prepare("SELECT run_key, snapshot_json, updated_at FROM run_snapshots WHERE account_id = ?")
    .bind(accountId)
    .all();

  const snapshotMap = new Map();
  for (const row of snapshotRows?.results || []) {
    try {
      snapshotMap.set(row.run_key, JSON.parse(row.snapshot_json));
    } catch {
      // skip bad snapshot
    }
  }

  const now = Date.now();
  let totalGoldEarned = 0;
  let totalPlayerXpEarned = 0;
  let totalMapFragmentsEarned = 0;
  let changedRuns = 0;
  const debugRuns = [];
  const notificationsToAdd = [];
  const snapshotStatements = [];
  const activeRunStatements = [];

  const nextInventory = { ...(profile.inventory || {}) };
  const nextBiomeProgress = { ...(profile.biomeProgress || {}) };

  for (const row of activeRunsRows?.results || []) {
    if (row.status !== "running") continue;

    const activeRun = {
      slotIndex: row.slot_index,
      runtimeKey: row.runtime_key,
      mapId: row.map_id,
      sourceMapId: row.source_map_id,
      isCrafted: Boolean(row.is_crafted),
      craftedMapId: row.crafted_map_id,
      status: row.status,
      startedAt: row.started_at,
      lastSimulatedAt: row.last_simulated_at,
      selected: Boolean(row.selected),
      opened: Boolean(row.opened),
      wave: row.wave,
      highestWaveReached: row.highest_wave_reached,
      updatedAt: row.updated_at,
    };

    const runtimeMap = getRuntimeMapFromActiveRun(activeRun, profile);
    if (!runtimeMap) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        skipped: "missing_runtime_map",
      });
      continue;
    }

    const hadStoredSnapshot = snapshotMap.has(activeRun.runtimeKey);
    const snapshot =
      snapshotMap.get(activeRun.runtimeKey) ||
      createFallbackSnapshotFromActiveRun(activeRun, runtimeMap, now);

    const safestLastSimulatedAt = Math.max(
      Number(activeRun.lastSimulatedAt || 0),
      Number(snapshot.lastSimulatedAt || 0)
    );

    if (!safestLastSimulatedAt) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        hadStoredSnapshot,
        skipped: "missing_last_simulated_at",
      });
      continue;
    }

    const missedMs = Math.max(0, now - safestLastSimulatedAt);
    if (missedMs <= 0) {
      debugRuns.push({
        runtimeKey: activeRun.runtimeKey,
        hadStoredSnapshot,
        safestLastSimulatedAt,
        now,
        missedMs,
        skipped: "non_positive_missed_ms",
      });
      continue;
    }

    const enemyIdStart =
      Math.max(1, ...((snapshot.enemies || []).map((enemy) => Number(enemy.id) || 0)), 0) + 1;

    const waveBefore = Number(snapshot.wave || activeRun.wave || 1);
    const enemiesBefore = Array.isArray(snapshot.enemies) ? snapshot.enemies.length : 0;
    const enemyPathPositionsBefore = Array.isArray(snapshot.enemies)
      ? snapshot.enemies.slice(0, 5).map((enemy) => Number(enemy?.pathPosition || 0))
      : [];

    const catchupResult = applyOfflineCatchupToRun({
      mapData: runtimeMap,
      runState: {
        ...snapshot,
        isRunning: true,
      },
      missedMs,
      playerProfile: profile,
      enemyIdStart,
    });

    const enemyPathPositionsAfter = Array.isArray(catchupResult?.nextRunState?.enemies)
      ? catchupResult.nextRunState.enemies.slice(0, 5).map((enemy) => Number(enemy?.pathPosition || 0))
      : [];

    debugRuns.push({
      runtimeKey: activeRun.runtimeKey,
      hadStoredSnapshot,
      safestLastSimulatedAt,
      now,
      missedMs,
      appliedMs: catchupResult.appliedMs || 0,
      waveBefore,
      waveAfter: Number(catchupResult?.nextRunState?.wave || waveBefore),
      enemiesBefore,
      enemiesAfter: Array.isArray(catchupResult?.nextRunState?.enemies) ? catchupResult.nextRunState.enemies.length : enemiesBefore,
      enemyPathPositionsBefore,
      enemyPathPositionsAfter,
      goldEarned: Number(catchupResult?.goldEarned || 0),
      xpEarned: Number(catchupResult?.playerXpEarned || 0),
      changed:
        Boolean(catchupResult.appliedMs) &&
        (
          Number(catchupResult?.nextRunState?.wave || waveBefore) !== waveBefore ||
          (Array.isArray(catchupResult?.nextRunState?.enemies) ? catchupResult.nextRunState.enemies.length : enemiesBefore) !== enemiesBefore ||
          JSON.stringify(enemyPathPositionsBefore) !== JSON.stringify(enemyPathPositionsAfter) ||
          Number(catchupResult?.goldEarned || 0) > 0
        ),
    });

    if (!catchupResult.appliedMs) continue;

    changedRuns += 1;
    totalGoldEarned += catchupResult.goldEarned || 0;
    totalPlayerXpEarned += catchupResult.playerXpEarned || 0;
    totalMapFragmentsEarned += catchupResult.mapFragmentsEarned || 0;

    const rareMaterialName = getRareMaterialNameForMap(runtimeMap);
    if ((catchupResult.rareDropsEarned || 0) > 0) {
      nextInventory[rareMaterialName] = (nextInventory[rareMaterialName] || 0) + (catchupResult.rareDropsEarned || 0);
    }

    if ((catchupResult.mapFragmentsEarned || 0) > 0) {
      nextInventory.mapFragments = (nextInventory.mapFragments || 0) + (catchupResult.mapFragmentsEarned || 0);
    }

    const biomeKey = getBiomeKeyFromMap(runtimeMap);
    const currentBiome = nextBiomeProgress[biomeKey] || { points: 0, dropBonusPct: 0 };
    const nextPoints = currentBiome.points + (catchupResult.biomeProgressEarned || 0);
    nextBiomeProgress[biomeKey] = {
      points: nextPoints,
      dropBonusPct: nextPoints / 10000,
    };

    notificationsToAdd.push(
      createOfflineCatchupNotification({
        mapName: runtimeMap.name,
        appliedMs: catchupResult.appliedMs,
        goldEarned: catchupResult.goldEarned || 0,
        xpEarned: catchupResult.playerXpEarned || 0,
        rareDrops: catchupResult.rareDropsEarned || 0,
        mapFragments: catchupResult.mapFragmentsEarned || 0,
        modifiers: 0,
      })
    );

    snapshotStatements.push(
      env.DB
        .prepare(
          `INSERT INTO run_snapshots (account_id, run_key, snapshot_json, updated_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(account_id, run_key) DO UPDATE SET
             snapshot_json = excluded.snapshot_json,
             updated_at = excluded.updated_at`
        )
        .bind(accountId, activeRun.runtimeKey, JSON.stringify(catchupResult.nextRunState), now)
    );

    activeRunStatements.push(
      env.DB
        .prepare(
          `UPDATE active_runs
           SET wave = ?, highest_wave_reached = ?, last_simulated_at = ?, updated_at = ?
           WHERE account_id = ? AND slot_index = ?`
        )
        .bind(
          Number(catchupResult.nextRunState.wave || activeRun.wave || 1),
          Number(catchupResult.nextRunState.highestWaveReached || activeRun.highestWaveReached || 1),
          now,
          now,
          accountId,
          Number(activeRun.slotIndex)
        )
    );
  }

  if (!changedRuns) {
    return json({
      ok: true,
      changedRuns: 0,
      totalGoldEarned: 0,
      totalPlayerXpEarned: 0,
      totalMapFragmentsEarned: 0,
      notificationsAdded: 0,
      debugRuns,
    });
  }

  const nextProfile = {
    ...profile,
    gold: (profile.gold || 0) + totalGoldEarned,
    inventory: nextInventory,
    biomeProgress: nextBiomeProgress,
    stats: {
      ...(profile.stats || {}),
      lifetimeGoldEarned: (profile.stats?.lifetimeGoldEarned || 0) + totalGoldEarned,
      highestWaveReached: Math.max(
        profile.stats?.highestWaveReached || 1,
        ...snapshotStatements.map(() => 1)
      ),
    },
  };

  const notificationsRow = await env.DB
    .prepare("SELECT notifications_json FROM notifications WHERE account_id = ?")
    .bind(accountId)
    .first();

  let existingNotifications = [];
  try {
    existingNotifications = JSON.parse(notificationsRow?.notifications_json || "[]");
    if (!Array.isArray(existingNotifications)) existingNotifications = [];
  } catch {
    existingNotifications = [];
  }

  const mergedNotifications = [...notificationsToAdd, ...existingNotifications]
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0))
    .slice(0, 25);

  const highestWaveReachedAfterCatchup = Math.max(
    profile.stats?.highestWaveReached || 1,
    ...((activeRunsRows?.results || []).map((row) => Number(row.highest_wave_reached || row.wave || 1)))
  );

  nextProfile.stats.highestWaveReached = highestWaveReachedAfterCatchup;

  const statements = [
    env.DB
      .prepare("UPDATE player_profiles SET profile_json = ?, updated_at = ? WHERE account_id = ?")
      .bind(JSON.stringify(nextProfile), now, accountId),
    env.DB
      .prepare(
        `INSERT INTO notifications (account_id, notifications_json, updated_at)
         VALUES (?, ?, ?)
         ON CONFLICT(account_id) DO UPDATE SET
           notifications_json = excluded.notifications_json,
           updated_at = excluded.updated_at`
      )
      .bind(accountId, JSON.stringify(mergedNotifications), now),
    ...snapshotStatements,
    ...activeRunStatements,
  ];

  await env.DB.batch(statements);

  return json({
    ok: true,
    changedRuns,
    totalGoldEarned,
    totalPlayerXpEarned,
    totalMapFragmentsEarned,
    notificationsAdded: notificationsToAdd.length,
    debugRuns,
  });
}


function getBossKeyForSummary(wave, runtimeMap) {
  const interval = Number(runtimeMap?.bossConfig?.interval || DEFAULT_BOSS_INTERVAL);
  if (!wave || wave < 1 || wave % interval !== 0) return null;
  const rotation = runtimeMap?.bossConfig?.rotation || [];
  if (!rotation.length) return null;
  const tier = Math.max(1, Math.floor(wave / interval));
  return rotation[(tier - 1) % rotation.length] || null;
}

function buildRunSummaryFromSnapshot(snapshot, runtimeMap) {
  const wave = Number(snapshot?.wave || 1);
  const highestWaveReached = Number(snapshot?.highestWaveReached || wave || 1);
  const enemies = Array.isArray(snapshot?.enemies) ? snapshot.enemies : [];
  const towers = Array.isArray(snapshot?.towers) ? snapshot.towers : [];
  const tracker = snapshot?.tracker || {};
  const elapsedMs = Math.max(1, Number(snapshot?.elapsedMs || 0));
  const bossEnemy = enemies.find((enemy) => enemy?.isBoss) || null;
  const interval = Number(runtimeMap?.bossConfig?.interval || DEFAULT_BOSS_INTERVAL);
  const nextBossWave = Math.ceil(Math.max(1, wave) / interval) * interval;
  const goldEarned = Number(tracker?.goldEarned || 0);
  const xpEarned = Number(tracker?.playerXpEarned || 0);
  const goldPerHour = elapsedMs > 0 ? (goldEarned * 3600000) / elapsedMs : 0;
  const xpPerHour = elapsedMs > 0 ? (xpEarned * 3600000) / elapsedMs : 0;
  const castleLives = Number(snapshot?.castleLives || runtimeMap?.maxLeaks || 20);
  const maxLeaks = Number(runtimeMap?.maxLeaks || 20);

  return {
    runtimeKey: snapshot?.mapId || runtimeMap?.id || null,
    mapId: runtimeMap?.id || null,
    mapName: runtimeMap?.name || "Unknown Map",
    wave,
    highestWaveReached,
    isRunning: Boolean(snapshot?.isRunning),
    enemiesCount: enemies.length,
    towersCount: towers.length,
    bossActive: Boolean(bossEnemy),
    activeBoss: bossEnemy
      ? {
          key: bossEnemy.bossType || bossEnemy.type || null,
          label: bossEnemy.label || "Boss",
          hp: Number(bossEnemy.hp || 0),
          maxHp: Number(bossEnemy.maxHp || 0),
          tier: Number(bossEnemy.bossTier || 1),
        }
      : null,
    currentBossKey: getBossKeyForSummary(wave, runtimeMap),
    nextBossWave,
    goldPerHour,
    xpPerHour,
    leaksUsed: Math.max(0, maxLeaks - castleLives),
    maxLeaks,
    status: snapshot?.status || "",
  };
}

async function handleGetRunSummary(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");
  const runtimeKey = url.searchParams.get("runtimeKey");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { runtimeMap, snapshot } = loaded;

  return json({
    ok: true,
    summary: buildRunSummaryFromSnapshot(snapshot, runtimeMap),
    serverTime: Date.now(),
  });
}

function calculateTowerDpsForSummary(tower) {
  const cooldownSeconds = Math.max(0.001, Number(tower?.cooldownMs || 1000) / 1000);
  const echoMultiplier = 1 + (Number(tower?.echoChance || 0) / 100);
  const directDps = (Number(tower?.damage || 0) / cooldownSeconds) * echoMultiplier;

  if (tower?.type === "cannon" && Number(tower?.splashRadius || 0) > 0 && Number(tower?.splashMultiplier || 0) > 0) {
    const splashDamage = Math.max(1, Math.floor(Number(tower?.damage || 0) * Number(tower?.splashMultiplier || 0)));
    return directDps + ((splashDamage / cooldownSeconds) * echoMultiplier);
  }

  return directDps;
}

function buildTowerMasterySummaryForSnapshotTower(tower, profile) {
  const loadout = getTowerLoadout(profile, tower?.type, tower?.masteryLoadoutId);
  const selectedIds = new Set(loadout?.selectedNodeIds || []);
  const purchasedIds = new Set(loadout?.purchasedNodeIds || []);
  const towerLevel = getTowerLevelFromXp(Number(tower?.xp || 0));

  return getTowerMasteryNodes(tower?.type).map((node) => ({
    id: node.id,
    name: node.name,
    size: node.size,
    specialist: Boolean(node.specialist),
    branch: node.branch,
    ...getTowerNodeState(node, selectedIds, purchasedIds, towerLevel),
  }));
}

function buildTowerSummaryPayload(tower, profile) {
  const loadout = getTowerLoadout(profile, tower?.type, tower?.masteryLoadoutId);

  return {
    id: Number(tower?.id || 0),
    type: tower?.type || null,
    label: TOWER_TYPES[tower?.type]?.label || tower?.type || "Tower",
    x: Number(tower?.x || 0),
    y: Number(tower?.y || 0),
    level: Number(tower?.level || getTowerLevelFromXp(Number(tower?.xp || 0))),
    xp: Number(tower?.xp || 0),
    damage: Number(tower?.damage || 0),
    range: Number(tower?.range || 0),
    cooldownMs: Number(tower?.cooldownMs || 0),
    dps: calculateTowerDpsForSummary(tower),
    kills: Number(tower?.kills || 0),
    lifetimeGoldEarned: Number(tower?.lifetimeGoldEarned || 0),
    targetPriority: tower?.targetPriority || "first",
    totalGoldSpent: Number(tower?.totalGoldSpent || tower?.baseCost || 0),
    refund: getTowerRefundValue(tower),
    upgradeCosts: {
      damage: getTowerUpgradeCost(tower, "damage"),
      range: getTowerUpgradeCost(tower, "range"),
      attackSpeed: getTowerUpgradeCost(tower, "attackSpeed"),
    },
    loadout: loadout
      ? {
          id: String(loadout.id),
          name: loadout.name || "Loadout",
        }
      : null,
    masterySummary: buildTowerMasterySummaryForSnapshotTower(tower, profile),
    modifier: tower?.modifier ? buildModifierPreview(tower.modifier) : null,
  };
}

async function handleGetTowerSummary(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");
  const runtimeKey = url.searchParams.get("runtimeKey");
  const towerId = Number(url.searchParams.get("towerId"));

  if (!accountId || !runtimeKey || !Number.isFinite(towerId)) {
    return json({ ok: false, error: "accountId, runtimeKey, and towerId are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { profile, snapshot } = loaded;
  const towers = Array.isArray(snapshot?.towers) ? snapshot.towers : [];
  const tower = towers.find((item) => Number(item?.id) === towerId);

  if (!tower) {
    return json({ ok: false, error: "Tower not found." }, { status: 404 });
  }

  return json({
    ok: true,
    summary: buildTowerSummaryPayload(tower, profile),
    serverTime: Date.now(),
  });
}

async function handleGetRunDetails(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");
  const runtimeKey = url.searchParams.get("runtimeKey");

  if (!accountId || !runtimeKey) {
    return json({ ok: false, error: "accountId and runtimeKey are required." }, { status: 400 });
  }

  const loaded = await loadProfileActiveRunSnapshotForRuntime(env, accountId, runtimeKey);
  if (loaded.error) return loaded.error;

  const { profile, runtimeMap, snapshot } = loaded;
  const towers = Array.isArray(snapshot?.towers) ? snapshot.towers : [];

  return json({
    ok: true,
    details: {
      runtimeKey,
      summary: buildRunSummaryFromSnapshot(snapshot, runtimeMap),
      towers: towers.map((tower) => buildTowerSummaryPayload(tower, profile)),
    },
    serverTime: Date.now(),
  });
}

async function handleGetGameplayState(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const catchupResult = await runBackendCatchupForAccount(env, accountId);
  if (!catchupResult?.ok) {
    return json({ ok: false, error: catchupResult?.error || "Catchup failed." }, { status: 400 });
  }

  const [profileRow, notificationsRow, mapSlotsRow, activeRunsRows, snapshotRows] = await Promise.all([
    env.DB
      .prepare("SELECT profile_json, updated_at FROM player_profiles WHERE account_id = ?")
      .bind(accountId)
      .first(),
    env.DB
      .prepare("SELECT notifications_json, updated_at FROM notifications WHERE account_id = ?")
      .bind(accountId)
      .first(),
    env.DB
      .prepare("SELECT map_slots_json, updated_at FROM map_slots WHERE account_id = ?")
      .bind(accountId)
      .first(),
    env.DB
      .prepare(
        `SELECT slot_index, runtime_key, map_id, source_map_id, is_crafted, crafted_map_id,
                status, started_at, last_simulated_at, selected, opened, wave,
                highest_wave_reached, updated_at
         FROM active_runs
         WHERE account_id = ?
         ORDER BY slot_index ASC`
      )
      .bind(accountId)
      .all(),
    env.DB
      .prepare("SELECT run_key, snapshot_json, updated_at FROM run_snapshots WHERE account_id = ?")
      .bind(accountId)
      .all(),
  ]);

  if (!profileRow?.profile_json) {
    return json({ ok: false, error: "Profile not found." }, { status: 404 });
  }

  let profile = null;
  try {
    profile = JSON.parse(profileRow.profile_json);
  } catch {
    return json({ ok: false, error: "Stored profile is invalid JSON." }, { status: 500 });
  }

  let notifications = [];
  try {
    notifications = JSON.parse(notificationsRow?.notifications_json || "[]");
    if (!Array.isArray(notifications)) notifications = [];
  } catch {
    notifications = [];
  }

  let mapSlots = [null, null, null];
  try {
    const parsed = JSON.parse(mapSlotsRow?.map_slots_json || "[null,null,null]");
    if (Array.isArray(parsed)) {
      mapSlots = parsed.slice(0, 3);
      while (mapSlots.length < 3) mapSlots.push(null);
    }
  } catch {
    mapSlots = [null, null, null];
  }

  const snapshots = {};
  for (const row of snapshotRows?.results || []) {
    try {
      snapshots[row.run_key] = JSON.parse(row.snapshot_json);
    } catch {
      // skip invalid row
    }
  }

  const activeRuns = (activeRunsRows?.results || []).map((row) => ({
    slotIndex: row.slot_index,
    runtimeKey: row.runtime_key,
    mapId: row.map_id,
    sourceMapId: row.source_map_id,
    isCrafted: Boolean(row.is_crafted),
    craftedMapId: row.crafted_map_id,
    status: row.status,
    startedAt: row.started_at,
    lastSimulatedAt: row.last_simulated_at,
    selected: Boolean(row.selected),
    opened: Boolean(row.opened),
    wave: row.wave,
    highestWaveReached: row.highest_wave_reached,
    updatedAt: row.updated_at,
  }));

  return json({
    ok: true,
    profile,
    profileUpdatedAt: profileRow.updated_at,
    notifications,
    notificationsUpdatedAt: notificationsRow?.updated_at ?? null,
    mapSlots,
    mapSlotsUpdatedAt: mapSlotsRow?.updated_at ?? null,
    activeRuns,
    snapshots,
    serverTime: Date.now(),
    catchup: {
      changedRuns: catchupResult.changedRuns || 0,
      notificationsAdded: catchupResult.notificationsAdded || 0,
    },
  });
}

async function handleDebugStatus(request, env) {
  const url = new URL(request.url);
  const accountId = url.searchParams.get("accountId");

  if (!accountId) {
    return json({ ok: false, error: "accountId is required." }, { status: 400 });
  }

  const accountRow = await env.DB
    .prepare("SELECT id, username, created_at FROM accounts WHERE id = ?")
    .bind(accountId)
    .first();

  if (!accountRow) {
    return json({ ok: false, error: "Account not found." }, { status: 404 });
  }

  const profileRow = await env.DB
    .prepare("SELECT updated_at, profile_json FROM player_profiles WHERE account_id = ?")
    .bind(accountId)
    .first();

  const notificationsRow = await env.DB
    .prepare("SELECT notifications_json, updated_at FROM notifications WHERE account_id = ?")
    .bind(accountId)
    .first();

  const mapSlotsRow = await env.DB
    .prepare("SELECT map_slots_json, updated_at FROM map_slots WHERE account_id = ?")
    .bind(accountId)
    .first();

  const activeRunsRows = await env.DB
    .prepare("SELECT slot_index, runtime_key, status, wave, highest_wave_reached, opened, selected, updated_at FROM active_runs WHERE account_id = ? ORDER BY slot_index ASC")
    .bind(accountId)
    .all();

  const snapshotRows = await env.DB
    .prepare("SELECT run_key, updated_at, snapshot_json FROM run_snapshots WHERE account_id = ? ORDER BY run_key ASC")
    .bind(accountId)
    .all();

  let profileSummary = null;
  if (profileRow?.profile_json) {
    try {
      const profile = JSON.parse(profileRow.profile_json);
      profileSummary = {
        profileVersion: profile?.profileVersion ?? null,
        name: profile?.name ?? null,
        level: profile?.level ?? null,
        gold: profile?.gold ?? null,
        hasCompletedSignup: Boolean(profile?.hasCompletedSignup),
        selectedMapId: profile?.appState?.selectedMapId ?? null,
        activeMapId: profile?.appState?.activeMapId ?? null,
        modifiersCount: Array.isArray(profile?.modifiers) ? profile.modifiers.length : 0,
        craftedMapsCount: Array.isArray(profile?.craftedMaps) ? profile.craftedMaps.length : 0,
      };
    } catch {
      profileSummary = {
        parseError: true,
      };
    }
  }

  const runSnapshots = (snapshotRows?.results || []).map((row) => {
    let parsed = null;
    try {
      parsed = JSON.parse(row.snapshot_json);
    } catch {
      parsed = null;
    }

    return {
      runKey: row.run_key,
      updatedAt: row.updated_at,
      wave: parsed?.wave ?? null,
      highestWaveReached: parsed?.highestWaveReached ?? null,
      mapId: parsed?.mapId ?? null,
      isRunning: parsed?.isRunning ?? null,
      towersCount: Array.isArray(parsed?.towers) ? parsed.towers.length : 0,
      enemiesCount: Array.isArray(parsed?.enemies) ? parsed.enemies.length : 0,
      elapsedMs: parsed?.elapsedMs ?? null,
    };
  });

  return json({
    ok: true,
    account: {
      id: accountRow.id,
      username: accountRow.username,
      createdAt: accountRow.created_at,
    },
    profile: {
      exists: Boolean(profileRow),
      updatedAt: profileRow?.updated_at ?? null,
      summary: profileSummary,
      jsonSizeBytes: profileRow?.profile_json ? new TextEncoder().encode(profileRow.profile_json).length : 0,
    },
    notifications: {
      exists: Boolean(notificationsRow),
      updatedAt: notificationsRow?.updated_at ?? null,
      count: (() => {
        try {
          const parsed = JSON.parse(notificationsRow?.notifications_json || "[]");
          return Array.isArray(parsed) ? parsed.length : 0;
        } catch {
          return 0;
        }
      })(),
      jsonSizeBytes: notificationsRow?.notifications_json
        ? new TextEncoder().encode(notificationsRow.notifications_json).length
        : 0,
    },
    mapSlots: {
      exists: Boolean(mapSlotsRow),
      updatedAt: mapSlotsRow?.updated_at ?? null,
      count: (() => {
        try {
          const parsed = JSON.parse(mapSlotsRow?.map_slots_json || "[null,null,null]");
          return Array.isArray(parsed) ? parsed.filter(Boolean).length : 0;
        } catch {
          return 0;
        }
      })(),
      jsonSizeBytes: mapSlotsRow?.map_slots_json
        ? new TextEncoder().encode(mapSlotsRow.map_slots_json).length
        : 0,
    },
    activeRuns: {
      count: (activeRunsRows?.results || []).length,
      items: (activeRunsRows?.results || []).map((row) => ({
        slotIndex: row.slot_index,
        runtimeKey: row.runtime_key,
        status: row.status,
        wave: row.wave,
        highestWaveReached: row.highest_wave_reached,
        opened: Boolean(row.opened),
        selected: Boolean(row.selected),
        updatedAt: row.updated_at,
      })),
    },
    runSnapshots: {
      count: runSnapshots.length,
      items: runSnapshots,
    },
    serverTime: Date.now(),
  });
}

export default {
  async fetch(request, env) {
    try {
      if (request.method === "OPTIONS") {
        return json({ ok: true });
      }

      if (url.pathname === "/api/admin/profile-action" && request.method === "POST") {
        return handleAdminProfileAction(request, env);
      }

      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/api/health") {
        return json({ ok: true, service: "idletd-backend" });
      }

      if (request.method === "GET" && url.pathname === "/api/game-config") {
        return handleGetGameConfig(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/signup") {
        return handleSignup(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/profile") {
        return handleGetProfile(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/profile/save") {
        return handleSaveProfile(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/run-snapshots") {
        return handleGetRunSnapshots(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-snapshots/save") {
        return handleSaveRunSnapshots(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/notifications") {
        return handleGetNotifications(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/notifications/save") {
        return handleSaveNotifications(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/map-slots") {
        return handleGetMapSlots(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/map-slots/save") {
        return handleSaveMapSlots(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/active-runs") {
        return handleGetActiveRuns(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/save") {
        return handleSaveActiveRuns(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/upsert") {
        return handleUpsertActiveRun(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/start") {
        return handleStartActiveRun(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/stop") {
        return handleStopActiveRun(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/reset") {
        return handleResetActiveRun(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/active-runs/catchup") {
        return handleCatchupActiveRuns(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/chat/global") {
        return handleGetGlobalChat(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/chat/global/send") {
        return handleSendGlobalChat(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/place-tower") {
        return handlePlaceTower(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/upgrade-tower") {
        return handleUpgradeTower(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/remove-tower") {
        return handleRemoveTower(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/set-target-priority") {
        return handleSetTowerPriority(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/equip-modifier") {
        return handleEquipModifier(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/unequip-modifier") {
        return handleUnequipModifier(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/admin/set-wave") {
        return handleAdminSetWave(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/admin/clear-enemies") {
        return handleAdminClearEnemies(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/admin/spawn-test-enemy") {
        return handleAdminSpawnTestEnemy(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/admin/spawn-boss") {
        return handleAdminSpawnBoss(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-actions/assign-loadout") {
        return handleAssignTowerLoadout(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/progression/permanent-upgrade") {
        return handlePermanentTowerUpgrade(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/tower-mastery/save-loadout") {
        return handleSaveTowerLoadout(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/tower-mastery/rename-loadout") {
        return handleRenameTowerLoadout(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/tower-mastery/delete-loadout") {
        return handleDeleteTowerLoadout(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/tower-mastery/load-draft") {
        return handleLoadTowerDraft(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/tower-mastery/toggle-node") {
        return handleToggleTowerDraftNode(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/modifiers/craft") {
        return handleCraftModifier(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/modifiers/reroll") {
        return handleRerollModifier(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/modifiers/enhance") {
        return handleEnhanceModifier(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/modifiers/salvage") {
        return handleSalvageModifier(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/modifiers/preview") {
        return handleGetModifierPreview(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/maps/craft") {
        return handleCraftMap(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/maps/crafted-preview") {
        return handleGetCraftedMapPreview(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/runs/summary") {
        return handleGetRunSummary(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/towers/summary") {
        return handleGetTowerSummary(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/runs/details") {
        return handleGetRunDetails(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/map-slots/assign-normal") {
        return handleAssignMapSlot(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/map-slots/assign-crafted") {
        return handleAssignCraftedMapSlot(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-session/open") {
        return handleOpenRunSession(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-session/close") {
        return handleCloseRunSession(request, env);
      }

      if (request.method === "POST" && url.pathname === "/api/run-session/select") {
        return handleSelectRunSession(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/gameplay-state") {
        return handleGetGameplayState(request, env);
      }

      if (request.method === "GET" && url.pathname === "/api/debug/status") {
        return handleDebugStatus(request, env);
      }

      return json({ ok: false, error: "Not found." }, { status: 404 });
    } catch (error) {
      return json(
        {
          ok: false,
          error: error instanceof Error ? error.message : "Internal server error",
        },
        { status: 500 }
      );
    }
  },
};
