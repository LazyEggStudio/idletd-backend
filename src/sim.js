export const MAX_OFFLINE_CATCHUP_MS = 30 * 60 * 1000;
export const OFFLINE_CATCHUP_STEP_MS = 250;
export const ENEMY_BASE_HP = 24;
export const ENEMY_BASE_REWARD = 8;
export const WAVE_GROWTH = 0.18;
export const WAVE_INTERVAL_MS = 1300;
export const WAVE_SIZE = 4;
export const RARE_DROP_CHANCE = 0.25;
export const MAP_FRAGMENT_DROP_CHANCE = 0.02;
export const BOSS_MAP_FRAGMENT_DROP_CHANCE = 0.5;
export const DEFAULT_BOSS_INTERVAL = 20;
export const IDLE_ENEMY_MOVE_SPEED = 1 / 6;
export const DARK_GREEN_ENEMY_MOVE_SPEED = 1 / 4;

export const ENEMY_TYPES = {
  pink_circle: {
    key: "pink_circle",
    label: "Pink Circle",
    speed: IDLE_ENEMY_MOVE_SPEED,
    size: 16,
    rewardScale: 1,
    hpScale: 1,
  },
  dark_green_circle: {
    key: "dark_green_circle",
    label: "Dark Green Circle",
    speed: DARK_GREEN_ENEMY_MOVE_SPEED,
    size: 12,
    rewardScale: 1.15,
    hpScale: 0.9,
  },
  light_purple_circle: {
    key: "light_purple_circle",
    label: "Light Purple Circle",
    speed: 1 / 8,
    size: 22,
    rewardScale: 1.45,
    hpScale: 2.0,
  },
};

export const BOSS_TYPES = {
  guardian: {
    key: "guardian",
    label: "Cave Guardian",
    speed: 1 / 10,
    size: 30,
    hpMultiplier: 14,
    rewardMultiplier: 10,
  },
  broodmother: {
    key: "broodmother",
    label: "Broodmother",
    speed: 1 / 9,
    size: 28,
    hpMultiplier: 12,
    rewardMultiplier: 11,
  },
  titan: {
    key: "titan",
    label: "Stone Titan",
    speed: 1 / 12,
    size: 34,
    hpMultiplier: 18,
    rewardMultiplier: 13,
  },
};

export const TOWER_TYPES = {
  archer: {
    key: "archer",
    label: "Archer",
    cost: 30,
    range: 2.25,
    damage: 8,
    cooldownMs: 700,
  },
  cannon: {
    key: "cannon",
    label: "Cannon",
    cost: 50,
    range: 1.75,
    damage: 15,
    cooldownMs: 1400,
  },
  frost: {
    key: "frost",
    label: "Frost Tower",
    cost: 45,
    range: 2.0,
    damage: 4,
    cooldownMs: 950,
    unlockLevel: 3,
    slowAmount: 0.65,
    slowDurationMs: 1800,
  },
};

export const SAMPLE_MAPS = [
  {
    id: "grasslands",
    name: "Grasslands",
    biome: "Grasslands",
    width: 10,
    height: 6,
    maxLeaks: 20,
    bossConfig: { rotation: ["guardian", "broodmother", "titan"], interval: DEFAULT_BOSS_INTERVAL },
    path: [
      { x: 1, y: 3 }, { x: 2, y: 3 }, { x: 3, y: 3 }, { x: 4, y: 3 },
      { x: 4, y: 2 }, { x: 4, y: 1 }, { x: 5, y: 1 }, { x: 6, y: 1 },
      { x: 6, y: 2 }, { x: 6, y: 3 }, { x: 6, y: 4 }, { x: 7, y: 4 }, { x: 8, y: 4 },
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
    path: [
      { x: 1, y: 1 }, { x: 2, y: 1 }, { x: 3, y: 1 }, { x: 4, y: 1 },
      { x: 4, y: 2 }, { x: 4, y: 3 }, { x: 5, y: 3 }, { x: 6, y: 3 },
      { x: 7, y: 3 }, { x: 7, y: 4 }, { x: 8, y: 4 }, { x: 9, y: 4 }, { x: 10, y: 4 },
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
    path: [
      { x: 1, y: 4 }, { x: 1, y: 3 }, { x: 1, y: 2 }, { x: 2, y: 2 },
      { x: 3, y: 2 }, { x: 4, y: 2 }, { x: 4, y: 3 }, { x: 4, y: 4 },
      { x: 5, y: 4 }, { x: 6, y: 4 }, { x: 7, y: 4 }, { x: 7, y: 3 }, { x: 7, y: 2 },
    ],
  },
];

function distance(a, b) {
  return Math.hypot(a.x - b.x, a.y - b.y);
}

function getPointAlongPath(path, pathPosition) {
  if (!path?.length) return { x: 0, y: 0, done: true };
  if (pathPosition >= path.length - 1) {
    const end = path[path.length - 1];
    return { x: end.x, y: end.y, done: true };
  }

  const baseIndex = Math.floor(pathPosition);
  const nextIndex = Math.min(baseIndex + 1, path.length - 1);
  const progress = pathPosition - baseIndex;
  const current = path[baseIndex];
  const next = path[nextIndex];

  return {
    x: current.x + (next.x - current.x) * progress,
    y: current.y + (next.y - current.y) * progress,
    done: false,
  };
}

export function getBiomeKeyFromMap(mapData) {
  switch ((mapData?.biome || "").toLowerCase()) {
    case "grasslands":
      return "grasslands";
    case "ember ridge":
      return "ember";
    case "tidelands":
      return "tide";
    default:
      return "grasslands";
  }
}

export function getRareMaterialNameForMap(mapData) {
  switch (getBiomeKeyFromMap(mapData)) {
    case "ember":
      return "Ember Shard";
    case "tide":
      return "Tideglass";
    default:
      return "Sunpetal";
  }
}

function getBiomeDropBonusPct(playerProfile, mapData) {
  const biomeKey = getBiomeKeyFromMap(mapData);
  return playerProfile?.biomeProgress?.[biomeKey]?.dropBonusPct || 0;
}

function getBiomeProgressGainForWave(wave) {
  return 1 + Math.floor((wave || 0) / 10);
}

function getCraftedRewardMultipliers(mapData) {
  return mapData?.craftedData?.rewardMultipliers || {
    gold: 1,
    xp: 1,
    rareDrop: 1,
    modifierRarity: 1,
  };
}

function getCraftedDangerModifiers(mapData) {
  return mapData?.craftedData?.dangerModifiers || {
    enemySpeed: 1,
    enemyHp: 1,
    bossInterval: mapData?.bossConfig?.interval || DEFAULT_BOSS_INTERVAL,
  };
}

function getBossTierForWave(wave, mapData) {
  const craftedDanger = getCraftedDangerModifiers(mapData);
  const interval = craftedDanger.bossInterval || mapData?.bossConfig?.interval || DEFAULT_BOSS_INTERVAL;
  return Math.max(1, Math.floor(wave / interval));
}

function getBossKeyForWave(wave, mapData) {
  const rotation = mapData?.bossConfig?.rotation || [];
  if (!rotation.length) return null;
  const tier = getBossTierForWave(wave, mapData);
  return rotation[(tier - 1) % rotation.length] || null;
}

function isBossWave(wave, mapData) {
  const craftedDanger = getCraftedDangerModifiers(mapData);
  const interval = craftedDanger.bossInterval || mapData?.bossConfig?.interval || DEFAULT_BOSS_INTERVAL;
  return wave > 0 && wave % interval === 0;
}

function getTowerTargetComparator(priority, towerPos) {
  if (priority === "closest") {
    return (a, b) => {
      const distA = distance(towerPos, a.point);
      const distB = distance(towerPos, b.point);
      if (distA !== distB) return distA - distB;
      return b.enemy.pathPosition - a.enemy.pathPosition;
    };
  }

  if (priority === "last") {
    return (a, b) => a.enemy.pathPosition - b.enemy.pathPosition;
  }

  return (a, b) => b.enemy.pathPosition - a.enemy.pathPosition;
}

function createBossEnemy({ wave, mapData, bossKey, enemyId }) {
  const bossType = BOSS_TYPES[bossKey];
  if (!bossType) return null;

  const craftedDanger = getCraftedDangerModifiers(mapData);
  const craftedRewards = getCraftedRewardMultipliers(mapData);
  const bossTier = getBossTierForWave(wave, mapData);
  const hpScale =
    (1 + (wave - 1) * WAVE_GROWTH) *
    bossType.hpMultiplier *
    (1 + (bossTier - 1) * 0.35) *
    (craftedDanger.enemyHp || 1);
  const rewardScale =
    (1 + (wave - 1) * 0.06) *
    bossType.rewardMultiplier *
    (1 + (bossTier - 1) * 0.2) *
    (craftedRewards.gold || 1);

  const bossHp = Math.max(1, Math.round(ENEMY_BASE_HP * hpScale));
  const bossReward = Math.max(1, Math.round(ENEMY_BASE_REWARD * rewardScale));

  return {
    id: enemyId,
    type: bossType.key,
    bossType: bossType.key,
    label: bossType.label,
    hp: bossHp,
    maxHp: bossHp,
    reward: bossReward,
    dropChance: 1,
    pathPosition: 0,
    speed: bossType.speed,
    size: bossType.size,
    slowedUntil: 0,
    slowMultiplier: 1,
    isBoss: true,
    bossTier,
    waveSpawned: wave,
  };
}

export function getRuntimeMapFromActiveRun(activeRun, profile) {
  if (!activeRun) return null;

  if (activeRun.isCrafted && activeRun.craftedMapId) {
    const craftedMap = (profile?.craftedMaps || []).find((item) => item.id === activeRun.craftedMapId);
    if (!craftedMap) return null;

    const baseMap =
      SAMPLE_MAPS.find((m) => m.id === craftedMap.sourceMapId) ||
      SAMPLE_MAPS.find((m) => m.biome === craftedMap.biome) ||
      SAMPLE_MAPS[0];

    return {
      ...baseMap,
      id: activeRun.runtimeKey || `crafted:${craftedMap.id}`,
      craftedMapId: craftedMap.id,
      craftedData: craftedMap,
      isCraftedMap: true,
      name: craftedMap.name,
      biome: craftedMap.biome,
      sourceMapId: craftedMap.sourceMapId,
    };
  }

  return SAMPLE_MAPS.find((m) => m.id === activeRun.sourceMapId || m.id === activeRun.mapId || m.id === activeRun.runtimeKey) || null;
}

export function createOfflineCatchupNotification({
  mapName,
  appliedMs,
  goldEarned = 0,
  xpEarned = 0,
  rareDrops = 0,
  mapFragments = 0,
  modifiers = 0,
}) {
  const parts = [];
  if (goldEarned > 0) parts.push(`+${goldEarned.toLocaleString()} gold`);
  if (xpEarned > 0) parts.push(`+${xpEarned.toLocaleString()} XP`);
  if (rareDrops > 0) parts.push(`+${rareDrops} rare drop${rareDrops === 1 ? "" : "s"}`);
  if (mapFragments > 0) parts.push(`+${mapFragments} map fragment${mapFragments === 1 ? "" : "s"}`);
  if (modifiers > 0) parts.push(`+${modifiers} modifier${modifiers === 1 ? "" : "s"}`);

  return {
    id: `offline-catchup-${mapName}-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
    type: "info",
    title: `Offline progress applied: ${Math.floor(appliedMs / 1000)}s on ${mapName}`,
    body: parts.length ? parts.join(" · ") : "No rewards earned during offline progress.",
    createdAt: Date.now(),
  };
}

export function simulateRunTick({
  mapData,
  runState,
  elapsedMs,
  deltaMs,
  spawnTimerMs,
  enemiesSpawnedInWave,
  enemyIdStart,
  playerProfile,
}) {
  const rareMaterialName = getRareMaterialNameForMap(mapData);
  const nextElapsedMs = elapsedMs + deltaMs;
  let nextSpawnTimerMs = spawnTimerMs + deltaMs;
  let nextEnemiesSpawnedInWave = enemiesSpawnedInWave;
  let nextEnemyId = enemyIdStart;
  const deltaSeconds = deltaMs / 1000;
  let goldEarned = 0;
  let rareDropsEarned = 0;
  let mapFragmentsEarned = 0;
  let biomeProgressEarned = 0;

  let nextState = {
    ...runState,
    highestWaveReached: Math.max(runState.highestWaveReached || 1, runState.wave || 1),
    enemies: (runState.enemies || []).map((enemy) => {
      const slowMultiplier =
        enemy.slowedUntil && enemy.slowedUntil > nextElapsedMs
          ? (enemy.slowMultiplier ?? 1)
          : 1;

      const craftedDanger = getCraftedDangerModifiers(mapData);
      const enemySpeedMultiplier = craftedDanger.enemySpeed || 1;

      return {
        ...enemy,
        pathPosition: enemy.pathPosition + enemy.speed * enemySpeedMultiplier * slowMultiplier * deltaSeconds,
      };
    }),
    towers: (runState.towers || []).map((tower) => ({ ...tower, lastShotAt: tower.lastShotAt || 0 })),
    popups: [],
    tracker: {
      goldEarned: runState?.tracker?.goldEarned || 0,
      playerXpEarned: runState?.tracker?.playerXpEarned || 0,
    },
    dropLog: Array.isArray(runState.dropLog) ? [...runState.dropLog] : [],
    killLog: Array.isArray(runState.killLog) ? [...runState.killLog] : [],
    leakLog: Array.isArray(runState.leakLog) ? [...runState.leakLog] : [],
    damageLog: Array.isArray(runState.damageLog) ? [...runState.damageLog] : [],
  };

  const currentWaveSize = WAVE_SIZE + Math.floor((nextState.wave - 1) * 1.5);

  if (nextEnemiesSpawnedInWave >= currentWaveSize) {
    if (nextState.enemies.length === 0) {
      const clearedWave = nextState.wave;
      biomeProgressEarned += getBiomeProgressGainForWave(clearedWave);
      nextEnemiesSpawnedInWave = 0;
      nextSpawnTimerMs = 0;
      nextState.wave += 1;
      nextState.highestWaveReached = Math.max(nextState.highestWaveReached || 1, nextState.wave);
      nextState.status = `Wave ${clearedWave} cleared.`;
    }
  } else if (nextSpawnTimerMs >= WAVE_INTERVAL_MS) {
    nextSpawnTimerMs = 0;

    if (isBossWave(nextState.wave, mapData)) {
      if (nextEnemiesSpawnedInWave === 0 && nextState.enemies.length === 0) {
        const bossKey = getBossKeyForWave(nextState.wave, mapData);
        const bossEnemy = createBossEnemy({
          wave: nextState.wave,
          mapData,
          bossKey,
          enemyId: nextEnemyId,
        });

        if (bossEnemy) {
          nextState.enemies.push(bossEnemy);
          nextEnemyId += 1;
          nextEnemiesSpawnedInWave = currentWaveSize;
          nextState.status = `Boss wave ${nextState.wave}: ${bossEnemy.label}`;
        }
      }
    } else {
      nextEnemiesSpawnedInWave += 1;

      let spawnType = ENEMY_TYPES.pink_circle;
      if (nextState.wave >= 5 && nextEnemiesSpawnedInWave % 5 === 0) {
        spawnType = ENEMY_TYPES.light_purple_circle;
      } else if (nextState.wave >= 3 && nextEnemiesSpawnedInWave % 3 === 0) {
        spawnType = ENEMY_TYPES.dark_green_circle;
      }

      const craftedDanger = getCraftedDangerModifiers(mapData);
      const craftedRewards = getCraftedRewardMultipliers(mapData);
      const hpScale =
        (1 + (nextState.wave - 1) * WAVE_GROWTH) *
        spawnType.hpScale *
        (craftedDanger.enemyHp || 1);
      const rewardScale =
        (1 + (nextState.wave - 1) * 0.06) *
        spawnType.rewardScale *
        (craftedRewards.gold || 1);

      const biomeBonusPct = getBiomeDropBonusPct(playerProfile, mapData);
      const baseDropChance = Math.min(0.5, RARE_DROP_CHANCE + (nextState.wave - 1) * 0.005);
      const finalRareDropMultiplier = (craftedRewards.rareDrop || 1) * (1 + biomeBonusPct / 100);
      const dropChance = Math.min(0.95, baseDropChance * finalRareDropMultiplier);

      nextState.enemies.push({
        id: nextEnemyId++,
        type: spawnType.key,
        label: spawnType.label,
        hp: Math.round(ENEMY_BASE_HP * hpScale),
        maxHp: Math.round(ENEMY_BASE_HP * hpScale),
        reward: Math.max(1, Math.round(ENEMY_BASE_REWARD * rewardScale)),
        dropChance,
        pathPosition: 0,
        speed: spawnType.speed,
        size: spawnType.size,
        slowedUntil: 0,
        slowMultiplier: 1,
        isBoss: false,
      });
    }
  }

  let leakedThisTick = 0;
  nextState.enemies = nextState.enemies.filter((enemy) => {
    const point = getPointAlongPath(mapData.path, enemy.pathPosition);
    if (point.done && enemy.pathPosition >= mapData.path.length - 1) {
      leakedThisTick += 1;
      return false;
    }
    return enemy.hp > 0;
  });

  if (leakedThisTick > 0) {
    nextState.castleLives = Math.max(0, nextState.castleLives - leakedThisTick);
    if (nextState.castleLives <= 0) {
      nextState = {
        ...nextState,
        wave: 1,
        castleLives: mapData.maxLeaks,
        enemies: [],
        popups: [],
        spawnTimerMs: 0,
        enemiesSpawnedInWave: 0,
        status: `Castle fell. Restarted ${mapData.name} at wave 1.`,
        startedAt: Date.now(),
      };

      return {
        nextState,
        nextElapsedMs: 0,
        nextSpawnTimerMs: WAVE_INTERVAL_MS - 50,
        nextEnemiesSpawnedInWave: 0,
        nextEnemyId,
        goldEarned,
        rareDropsEarned,
        mapFragmentsEarned,
        biomeProgressEarned,
      };
    }
  }

  if (nextState.towers.length > 0 && nextState.enemies.length > 0) {
    const killedEnemyIds = new Set();

    nextState.towers.forEach((tower) => {
      const cooldownDone = nextElapsedMs - (tower.lastShotAt || 0) >= (tower.cooldownMs || 1000);
      if (!cooldownDone) return;

      const towerPos = { x: tower.x, y: tower.y };
      const targets = nextState.enemies
        .filter((enemy) => !killedEnemyIds.has(enemy.id))
        .map((enemy) => ({ enemy, point: getPointAlongPath(mapData.path, enemy.pathPosition) }))
        .filter(({ point }) => distance(towerPos, point) <= (tower.range || 0))
        .sort(getTowerTargetComparator(tower.targetPriority || "first", towerPos));

      if (!targets.length) return;

      const target = targets[0].enemy;
      const towerDamage = tower.damage || 0;
      target.hp -= towerDamage;
      tower.lastShotAt = nextElapsedMs;

      if (tower.type === "frost") {
        target.slowedUntil = nextElapsedMs + (tower.slowDurationMs || 1800);
        target.slowMultiplier = tower.slowAmount || 0.65;
      }

      if (target.hp <= 0 && !killedEnemyIds.has(target.id)) {
        killedEnemyIds.add(target.id);
        goldEarned += target.reward || 0;

        nextState.tracker = {
          ...(nextState.tracker || { goldEarned: 0, playerXpEarned: 0 }),
          goldEarned: (nextState.tracker?.goldEarned || 0) + (target.reward || 0),
          playerXpEarned:
            (nextState.tracker?.playerXpEarned || 0) +
            Math.max(0, Math.floor((target.reward || 0) * 0.25 * (getCraftedRewardMultipliers(mapData).xp || 1))),
        };

        if (Math.random() < (target.dropChance ?? RARE_DROP_CHANCE)) {
          rareDropsEarned += 1;
        }

        const fragmentChance = target.isBoss ? BOSS_MAP_FRAGMENT_DROP_CHANCE : MAP_FRAGMENT_DROP_CHANCE;
        if (Math.random() < fragmentChance) {
          mapFragmentsEarned += 1;
        }
      }
    });

    nextState.enemies = nextState.enemies.filter((enemy) => !killedEnemyIds.has(enemy.id));
  }

  if (rareDropsEarned > 0) {
    nextState.drops = {
      ...(nextState.drops || {}),
      [rareMaterialName]: (nextState.drops?.[rareMaterialName] || 0) + rareDropsEarned,
    };
  }

  return {
    nextState,
    nextElapsedMs,
    nextSpawnTimerMs,
    nextEnemiesSpawnedInWave,
    nextEnemyId,
    goldEarned,
    rareDropsEarned,
    mapFragmentsEarned,
    biomeProgressEarned,
  };
}

export function applyOfflineCatchupToRun({
  mapData,
  runState,
  missedMs,
  playerProfile,
  enemyIdStart,
}) {
  const appliedMs = Math.max(0, Math.min(MAX_OFFLINE_CATCHUP_MS, missedMs || 0));
  if (!appliedMs || !runState?.isRunning) {
    return {
      appliedMs: 0,
      nextRunState: runState,
      nextEnemyId: enemyIdStart,
      goldEarned: 0,
      playerXpEarned: 0,
      rareDropsEarned: 0,
      mapFragmentsEarned: 0,
      earnedModifiers: [],
      biomeProgressEarned: 0,
    };
  }

  let workingRunState = {
    ...runState,
    elapsedMs: runState.elapsedMs || 0,
    spawnTimerMs: runState.spawnTimerMs || 0,
    enemiesSpawnedInWave: runState.enemiesSpawnedInWave || 0,
    enemies: (runState.enemies || []).filter((enemy) => enemy && enemy.hp > 0 && typeof enemy.pathPosition === "number"),
    towers: (runState.towers || []).map((tower) => ({ ...tower, lastShotAt: tower.lastShotAt || 0 })),
    popups: [],
  };

  let workingElapsedMs = workingRunState.elapsedMs || 0;
  let workingSpawnTimerMs = workingRunState.spawnTimerMs || 0;
  let workingEnemiesSpawned = workingRunState.enemiesSpawnedInWave || 0;
  let workingEnemyId = enemyIdStart || Math.max(1, ...(workingRunState.enemies || []).map((enemy) => Number(enemy.id) || 0)) + 1;

  let totalGoldEarned = 0;
  let totalRareDropsEarned = 0;
  let totalMapFragmentsEarned = 0;
  let totalPlayerXpEarned = 0;
  let totalBiomeProgressEarned = 0;

  const basePlayerXpBefore = workingRunState?.tracker?.playerXpEarned || 0;

  let remainingMs = appliedMs;
  while (remainingMs > 0) {
    const stepDeltaMs = Math.min(remainingMs, OFFLINE_CATCHUP_STEP_MS);

    const result = simulateRunTick({
      mapData,
      runState: workingRunState,
      elapsedMs: workingElapsedMs,
      deltaMs: stepDeltaMs,
      spawnTimerMs: workingSpawnTimerMs,
      enemiesSpawnedInWave: workingEnemiesSpawned,
      enemyIdStart: workingEnemyId,
      playerProfile,
    });

    workingRunState = result.nextState;
    workingElapsedMs = result.nextElapsedMs;
    workingSpawnTimerMs = result.nextSpawnTimerMs;
    workingEnemiesSpawned = result.nextEnemiesSpawnedInWave;
    workingEnemyId = result.nextEnemyId;

    totalGoldEarned += result.goldEarned || 0;
    totalRareDropsEarned += result.rareDropsEarned || 0;
    totalMapFragmentsEarned += result.mapFragmentsEarned || 0;
    totalBiomeProgressEarned += result.biomeProgressEarned || 0;

    remainingMs -= stepDeltaMs;
  }

  totalPlayerXpEarned = Math.max(
    0,
    (workingRunState?.tracker?.playerXpEarned || 0) - basePlayerXpBefore
  );

  return {
    appliedMs,
    nextRunState: {
      ...workingRunState,
      elapsedMs: workingElapsedMs,
      spawnTimerMs: workingSpawnTimerMs,
      enemiesSpawnedInWave: workingEnemiesSpawned,
      popups: [],
      lastSimulatedAt: Date.now(),
    },
    nextEnemyId: workingEnemyId,
    goldEarned: totalGoldEarned,
    playerXpEarned: totalPlayerXpEarned,
    rareDropsEarned: totalRareDropsEarned,
    mapFragmentsEarned: totalMapFragmentsEarned,
    earnedModifiers: [],
    biomeProgressEarned: totalBiomeProgressEarned,
  };
}
