export type BatchMode = "dual" | "standard" | "ludicrous";
export type AccountKey = "acc1" | "acc2" | "acc3";

const ACCOUNT_KEYS: AccountKey[] = ["acc1", "acc2", "acc3"];

const MODE_DEFAULT_BUDGET: Record<BatchMode, number> = {
  dual: 520,
  standard: 950,
  ludicrous: 600,
};

const MODE_TARGET_SECONDS: Record<BatchMode, number> = {
  dual: 18,
  standard: 9,
  ludicrous: 16,
};

const MODE_MIN_WORDS_PER_ACCOUNT: Record<BatchMode, number> = {
  dual: 280,
  standard: 500,
  ludicrous: 300,
};

const MODE_COORDINATION_PENALTY_SECONDS: Record<BatchMode, number> = {
  dual: 1.2,
  standard: 0.7,
  ludicrous: 1.0,
};

interface SchedulerAccountRates {
  health?: string;
  successRateDual?: number;
  successRateStandard?: number;
  successRateLudicrous?: number;
  retryRateDual?: number;
  retryRateStandard?: number;
  retryRateLudicrous?: number;
  timeoutRateDual?: number;
  timeoutRateStandard?: number;
  timeoutRateLudicrous?: number;
}

interface SchedulerRecommendedPerAccount {
  dual?: number;
  standard?: number;
  ludicrous?: number;
}

interface SchedulerRecommendedBudgets {
  dual?: number;
  standard?: number;
  ludicrous?: number;
  perAccount?: Partial<Record<AccountKey, SchedulerRecommendedPerAccount>>;
}

interface SchedulerRolling {
  successRatio?: number;
  fallbackRate?: number;
}

interface SchedulerSnapshot {
  recommendedBudgets?: SchedulerRecommendedBudgets;
  accounts?: Partial<Record<AccountKey, SchedulerAccountRates>>;
  rolling?: SchedulerRolling;
}

interface HealthAccountStatus {
  status?: string;
}

export interface HealthSnapshot {
  acc1?: HealthAccountStatus;
  acc2?: HealthAccountStatus;
  acc3?: HealthAccountStatus;
  scheduler?: SchedulerSnapshot;
}

type AccountBudgetProfile = {
  accountKey: AccountKey;
  rawBudget: number;
  effectiveBudget: number;
};

type CandidatePlan = {
  count: number;
  estimatedSeconds: number;
  effectiveCapacity: number;
  accounts: AccountKey[];
};

export interface AccountSelectionResult {
  count: number;
  estimatedSeconds: number;
  effectiveCapacity: number;
  accounts: AccountKey[];
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function toFiniteNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function modeSuffix(mode: BatchMode): "Dual" | "Standard" | "Ludicrous" {
  switch (mode) {
    case "dual":
      return "Dual";
    case "standard":
      return "Standard";
    case "ludicrous":
      return "Ludicrous";
  }
}

function isAccountReady(snapshot: HealthSnapshot | undefined, accountKey: AccountKey): boolean {
  const status = snapshot?.[accountKey]?.status?.toLowerCase() ?? "ready";
  return status === "ready" || status === "ok";
}

function isAccountTripped(snapshot: HealthSnapshot | undefined, accountKey: AccountKey): boolean {
  const health = snapshot?.scheduler?.accounts?.[accountKey]?.health?.toLowerCase();
  return health === "tripped";
}

function getAvailableAccounts(snapshot?: HealthSnapshot): AccountKey[] {
  const readyAccounts = ACCOUNT_KEYS.filter((accountKey) =>
    isAccountReady(snapshot, accountKey),
  );
  if (readyAccounts.length === 0) {
    return ["acc1"];
  }

  const healthyAccounts = readyAccounts.filter(
    (accountKey) => !isAccountTripped(snapshot, accountKey),
  );
  if (healthyAccounts.length > 0) {
    return healthyAccounts;
  }

  // If everything ready is currently marked tripped, keep at least one route.
  return [readyAccounts[0]];
}

function getRawBudget(
  snapshot: HealthSnapshot | undefined,
  accountKey: AccountKey,
  mode: BatchMode,
): number {
  const perAccountBudget = snapshot?.scheduler?.recommendedBudgets?.perAccount?.[
    accountKey
  ]?.[mode];
  const globalBudget = snapshot?.scheduler?.recommendedBudgets?.[mode];
  const resolved = toFiniteNumber(perAccountBudget) ?? toFiniteNumber(globalBudget);
  return clamp(resolved ?? MODE_DEFAULT_BUDGET[mode], 120, 2000);
}

function getReliabilityFactor(
  snapshot: HealthSnapshot | undefined,
  accountKey: AccountKey,
  mode: BatchMode,
): number {
  const account = snapshot?.scheduler?.accounts?.[accountKey];
  if (!account) {
    return 1;
  }

  const suffix = modeSuffix(mode);
  const successRate =
    toFiniteNumber(
      account[`successRate${suffix}` as keyof SchedulerAccountRates],
    ) ?? 1;
  const retryRate =
    toFiniteNumber(account[`retryRate${suffix}` as keyof SchedulerAccountRates]) ?? 0;
  const timeoutRate =
    toFiniteNumber(account[`timeoutRate${suffix}` as keyof SchedulerAccountRates]) ?? 0;

  const rateFactor = clamp(
    successRate - retryRate * 0.45 - timeoutRate * 0.8,
    0.4,
    1.05,
  );

  const health = account.health?.toLowerCase();
  const healthFactor =
    health === "degraded" ? 0.8 : health === "tripped" ? 0.35 : 1;

  return clamp(rateFactor * healthFactor, 0.35, 1.05);
}

function getSystemPenaltySeconds(
  snapshot: HealthSnapshot | undefined,
  accountCount: number,
): number {
  if (!snapshot?.scheduler?.rolling || accountCount <= 1) {
    return 0;
  }

  const successRatio = clamp(
    toFiniteNumber(snapshot.scheduler.rolling.successRatio) ?? 1,
    0,
    1,
  );
  const fallbackRate = clamp(
    toFiniteNumber(snapshot.scheduler.rolling.fallbackRate) ?? 0,
    0,
    1,
  );

  // Penalize extra parallelism when recent reliability is shaky.
  return (fallbackRate * 4 + (1 - successRatio) * 6) * (accountCount - 1);
}

function buildCandidatePlans(
  totalWords: number,
  mode: BatchMode,
  profiles: AccountBudgetProfile[],
  snapshot?: HealthSnapshot,
): CandidatePlan[] {
  const targetSeconds = MODE_TARGET_SECONDS[mode];
  const coordinationPenalty = MODE_COORDINATION_PENALTY_SECONDS[mode];
  const minWordsPerAccount = MODE_MIN_WORDS_PER_ACCOUNT[mode];
  const maxCountByWords = Math.max(
    1,
    Math.floor(totalWords / minWordsPerAccount),
  );
  const maxCandidateCount = Math.min(profiles.length, maxCountByWords);
  const candidates: CandidatePlan[] = [];

  for (let count = 1; count <= maxCandidateCount; count++) {
    const selected = profiles.slice(0, count);
    const capacity = selected.reduce((sum, profile) => sum + profile.effectiveBudget, 0);
    const safeCapacity = Math.max(capacity, 100);

    let estimatedSeconds = (totalWords / safeCapacity) * targetSeconds;
    estimatedSeconds += (count - 1) * coordinationPenalty;
    estimatedSeconds += getSystemPenaltySeconds(snapshot, count);

    candidates.push({
      count,
      estimatedSeconds,
      effectiveCapacity: safeCapacity,
      accounts: selected.map((profile) => profile.accountKey),
    });
  }

  return candidates;
}

export function chooseAdaptiveAccountCount(
  totalWords: number,
  mode: BatchMode,
  snapshot?: HealthSnapshot,
): AccountSelectionResult {
  if (!Number.isFinite(totalWords) || totalWords <= 0) {
    return {
      count: 1,
      estimatedSeconds: 0,
      effectiveCapacity: MODE_DEFAULT_BUDGET[mode],
      accounts: ["acc1"],
    };
  }

  const availableAccounts = getAvailableAccounts(snapshot);
  const profiles = availableAccounts
    .map((accountKey) => {
      const rawBudget = getRawBudget(snapshot, accountKey, mode);
      const reliabilityFactor = getReliabilityFactor(snapshot, accountKey, mode);
      return {
        accountKey,
        rawBudget,
        effectiveBudget: Math.max(120, rawBudget * reliabilityFactor),
      } satisfies AccountBudgetProfile;
    })
    .sort((a, b) => b.effectiveBudget - a.effectiveBudget);

  const candidates = buildCandidatePlans(totalWords, mode, profiles, snapshot);
  if (candidates.length === 0) {
    return {
      count: 1,
      estimatedSeconds: (totalWords / MODE_DEFAULT_BUDGET[mode]) * MODE_TARGET_SECONDS[mode],
      effectiveCapacity: MODE_DEFAULT_BUDGET[mode],
      accounts: ["acc1"],
    };
  }

  let best = candidates[0];
  // Only scale up when improvement is meaningful to avoid over-sharding tiny batches.
  for (let i = 1; i < candidates.length; i++) {
    const candidate = candidates[i];
    if (candidate.estimatedSeconds + 0.9 < best.estimatedSeconds) {
      best = candidate;
    }
  }

  return best;
}
