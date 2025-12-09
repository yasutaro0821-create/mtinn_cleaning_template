/**
 * mt.inn 15日シフト自動生成メインスクリプト (README が唯一仕様)
 *
 * 機能:
 * 1. Spreadsheet から StaffDB / Availability / Demand / ShiftPattern / Shift_15a / Line_* を取得
 * 2. 各フェーズ (Night → Cook → Bar+Dinner → Breakfast Core → Lobby AM → Lobby PM → Dinner → Breakfast → Cleaning → Part) を順に実行
 * 3. need / have / 労働時間 / 公休 / 連勤 を検算し、Shift_15a と Line_* に反映
 *
 * NOTE: 現時点ではフェーズ内の割当ロジックは骨組みのみ。段階的に実装を追加していく。
 */

const SPREADSHEET_ID = '1nCextT2bzlH44hO6VDrci5zxhv3_ZlCYAPx4k_AnWAQ';
const DAY_COUNT = 15;
const HOURS_SEQUENCE = buildHourSequence_();

/**
 * スプレッドシートを開いたときにメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('シフト生成')
    .addItem('シフトを生成', 'generateShiftPlan')
    .addToUi();
}

const CONFIG = {
  sheets: {
    staffDB: 'StaffDB',
    availability: 'Availability',
    demand: 'Demand_15',
    patterns: 'ShiftPattern',
    output: 'Shift_15a',
    linePrefix: 'Line_'
  },
  colHints: {
    name: ['氏名', '名前', 'Name'],
    division: ['区分', 'Division'],
    min: ['半月_MinHours', 'Min', 'MinHours'],
    max: ['半月_MaxHours', 'Max', 'MaxHours'],
    fulltimeFlag: ['フルタイムフラグ', 'FullTime'],
    staffId: ['StaffID', 'ID']
  },
  shiftPatternColHints: {
    code: ['コード', 'Code', 'ShiftCode', 'シフトコード'],
    department: ['部門', '部門名', 'Department'],
    hours: ['労働時間', 'hours', 'Hours', '時間', 'HourRange']
  },
  literals: {
    holiday: '公休',
    off: 'OFF',
    ng: 'NG'
  },
  colors: {
    warning: '#fff2cc',
    cleared: '#ffffff',
    needColumn: '#d9ead3',
    haveColumn: '#ead1dc',
    normal: '#ffffff',
    overMax: '#ffcccc',
    underMin: '#fff2cc',
    holiday: '#d9d9d9', // 公休（薄いグレー）
    off: '#b3b3b3', // OFF（中程度のグレー）
    ng: '#999999', // NG（濃いグレー）
    night: '#ffcccc', // ナイト（薄い赤）
    cook: '#ccffcc' // 調理（薄い緑）
  },
  limits: {
    maxConsecutive: 5, // 正社員・パート共通で最大5日
    defaultHolidayCount: 4
  },
  priority: {
    lobbyAmCore: ['村松', '三本木'],
    barCodes: ['FDB-S'],
    nightCodes: ['FN'],
    cookCodes: ['FC', 'FC-S'],
    breakfastCoreCodes: ['FBH', 'FB-S'],
    lobbyAmLongShortPairs: [
      ['FLDL-L', 'FL-S'],
      ['FL-L', 'FL-S']
    ],
    cookCombinations: [
      { members: ['国島', '山内'], weekdayOnly: false },
      { members: ['国島', '直美'], weekdayOnly: false },
      { members: ['山内', '直美'], weekdayOnly: true }
    ],
    cookStaffNames: {
      naomi: '直美',
      kunishima: '国島',
      yamauchi: '山内'
    },
    offHoursReduction: 8
  }
};

/**
 * エントリポイント: 15日分シフトを再生成（制約が完全一致するまでループ）
 */
function generateShiftPlan() {
  Logger.log('🚀 generateShiftPlan 開始');
  Logger.log(`対象スプレッドシートID: ${SPREADSHEET_ID}`);
  try {
    const ctx = loadContext_();
    Logger.log(`✅ コンテキスト読み込み完了: スタッフ${ctx.staff.length}名, パターン${Object.keys(ctx.patterns).length}個`);
    Logger.log(`実際に開いたスプレッドシートID: ${ctx.spreadsheet.getId()}`);
    Logger.log(`IDが一致しているか: ${ctx.spreadsheet.getId() === SPREADSHEET_ID}`);
    
    runPhases_(ctx);
    Logger.log('✅ フェーズ実行完了');
    
    finalizeAssignments_(ctx);
    Logger.log('✅ 割り当て確定完了');
    
    // 制約チェック
    const violations = checkAllConstraints_(ctx);
    if (violations.length > 0) {
      ctx.warnings.push(`制約違反 ${violations.length}件: ${violations.slice(0, 5).join(', ')}`);
      Logger.log(`⚠️ 制約違反が残っています: ${violations.length}件`);
    } else {
      Logger.log('✅ すべての制約を満たしました');
    }
    
    // デバッグ: 国島・直美・山内のassignmentsを確認
    const cookNames = CONFIG.priority.cookStaffNames;
    ['国島', '直美', '山内'].forEach(name => {
      const assignment = ctx.assignments[name];
      if (assignment) {
        const codes = assignment.days.map((code, idx) => `${idx + 1}日=${code || '(空)'}`).join(', ');
        Logger.log(`${name} の割り当て: ${codes}`);
      }
    });
    
    writeShiftSheet_(ctx);
    Logger.log('✅ Shift_15a 書き込み完了');
    
    writeLineSheets_(ctx);
    Logger.log('✅ Line_* 書き込み完了');
    
    Logger.log('✅ generateShiftPlan 完了');
  } catch (error) {
    Logger.log(`❌ エラー発生: ${error.toString()}`);
    Logger.log(`スタック: ${error.stack}`);
    throw error;
  }
}

/**
 * 必要なすべてのデータを読み込み、コンテキストを組み立てる
 * @returns {ShiftContext}
 */
function loadContext_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staff = readStaffDb_(spreadsheet);
  const availability = readAvailability_(spreadsheet);
  const patterns = readShiftPatterns_(spreadsheet);
  const demand = readDemand_(spreadsheet);

  const ctx = {
    spreadsheet,
    staff,
    availability,
    patterns,
    demand,
    assignments: initAssignments_(staff),
    metrics: initMetrics_(staff),
    logs: [],
    warnings: []
  };

  ctx.logs.push(`[load] staff=${staff.length}, availability=${Object.keys(availability).length}人, patterns=${Object.keys(patterns).length}種類`);
  return ctx;
}

/**
 * フェーズを定義順で実行
 * @param {ShiftContext} ctx
 */
function runPhases_(ctx) {
  // 重要: Availability の NG/OFF を先に配置（公休より先）
  assignAvailabilityFixed_(ctx);

  // ナイトと調理だけに特化
  const phases = [
    { name: 'NIGHT', fn: phaseNight_ },
    { name: 'COOK', fn: phaseCook_ }
  ];

  phases.forEach(phase => {
    const t0 = new Date();
    ctx.logs.push(`--- Phase ${phase.name} start ---`);
    phase.fn(ctx);
    const elapsed = ((new Date()) - t0) / 1000;
    ctx.logs.push(`--- Phase ${phase.name} end (${elapsed.toFixed(2)}s) ---`);
  });
}

/**
 * Availability の NG/OFF を先に配置（公休より先に実行）
 * @param {ShiftContext} ctx
 */
function assignAvailabilityFixed_(ctx) {
  Object.keys(ctx.availability).forEach(name => {
    const avail = ctx.availability[name];
    for (let day = 1; day <= DAY_COUNT; day++) {
      const val = avail[day];
      if (val === CONFIG.literals.ng || val === CONFIG.literals.off || val === CONFIG.literals.holiday) {
        ctx.assignments[name].days[day - 1] = val;
      }
    }
  });
  ctx.logs.push('[Availability] NG/OFF/公休 を先に配置しました');
}

/**
 * フェーズ別割当が終わった後の検算
 * @param {ShiftContext} ctx
 */
function finalizeAssignments_(ctx) {
  recalcMetrics_(ctx);
  validateNeedCoverage_(ctx);
  validateHourAndHoliday_(ctx);
}

/**
 * Night フェーズ（15日を4等分に近づけて公休を配置、残りに勤務を入れる）
 * @param {ShiftContext} ctx
 */
function phaseNight_(ctx) {
  const nightCode = CONFIG.priority.nightCodes.find(code => ctx.patterns[code]);
  if (!nightCode) {
    ctx.warnings.push('Night パターンが見つからないため Phase NIGHT をスキップしました');
    return;
  }
  const nightPattern = ctx.patterns[nightCode];
  const fullTimers = ctx.staff.filter(st => st.isFullTime && st.shiftFlags[nightCode]);
  const partTimers = ctx.staff.filter(st => !st.isFullTime && st.shiftFlags[nightCode]);

  if (!fullTimers.length) {
    ctx.warnings.push('正社員の Night 候補が存在しません');
    return;
  }

  const fullTimer = fullTimers[0]; // 長谷川さん（FN可否TRUEの正社員）
  const partTimer = partTimers.length > 0 ? partTimers[0] : null; // パートナイトは一人

  // ステップ1: 公休を15日を4等分に近づけて配置（3-4-4-4日間隔）
  const holidayDays = distributeHolidaysEvenly_(fullTimer, ctx);
  holidayDays.forEach(day => {
    if (!ctx.availability[fullTimer.name]?.[day]) {
      ctx.assignments[fullTimer.name].days[day - 1] = CONFIG.literals.holiday;
    }
  });

  // ステップ2: 残りの日に勤務を入れる（OFFがある場合は10回、ない場合は11回）
  for (let day = 1; day <= DAY_COUNT; day++) {
    const current = ctx.assignments[fullTimer.name].days[day - 1];
    const isHoliday = current === CONFIG.literals.holiday || current === CONFIG.literals.ng;
    const isOff = current === CONFIG.literals.off;
    
    // 公休でもOFFでもなく、needがある場合
    if (!isHoliday && !isOff && hasRemainingDemandForPattern_(ctx, day, nightPattern)) {
      if (canAssignPattern_(ctx, fullTimer, day, nightCode, { allowNightMinShortage: true })) {
        commitAssignment_(ctx, fullTimer, day, nightCode);
      }
    }
  }

  // ステップ3: 長谷川さんが入っていない日にパートナイトを入れる
  if (partTimer) {
    for (let day = 1; day <= DAY_COUNT; day++) {
      if (hasRemainingDemandForPattern_(ctx, day, nightPattern)) {
        const fullTimerCode = ctx.assignments[fullTimer.name].days[day - 1];
        const isFullTimerWorking = fullTimerCode && 
                                   fullTimerCode !== CONFIG.literals.holiday && 
                                   fullTimerCode !== CONFIG.literals.off && 
                                   fullTimerCode !== CONFIG.literals.ng &&
                                   fullTimerCode !== '';
        
        // 長谷川さんが入っていない場合、パートを入れる
        if (!isFullTimerWorking) {
          if (canAssignPattern_(ctx, partTimer, day, nightCode, { allowNightMinShortage: true })) {
            commitAssignment_(ctx, partTimer, day, nightCode);
            ctx.logs.push(`[Night] day${day} に ${partTimer.name} (パート) を配置`);
          }
        }
      }
    }
  }

  // 最終チェック
  if (hasUnmetNightNeed_(ctx, nightPattern)) {
    ctx.warnings.push('Night need をすべて満たせませんでした');
  }
}

/**
 * 公休を15日を4等分に近づけて配置（3-4-4-4日間隔）
 */
function distributeHolidaysEvenly_(staff, ctx) {
  const targetHolidayCount = CONFIG.limits.defaultHolidayCount; // 4日
  const holidays = [];
  
  // 既にAvailabilityで指定されている公休/NGを確認
  for (let day = 1; day <= DAY_COUNT; day++) {
    const avail = ctx.availability[staff.name]?.[day];
    if (avail === CONFIG.literals.holiday || avail === CONFIG.literals.ng) {
      holidays.push(day);
    }
  }
  
  // 既存の公休数が4未満の場合、均等に配置
  if (holidays.length < targetHolidayCount) {
    const needed = targetHolidayCount - holidays.length;
    const availableDays = [];
    for (let day = 1; day <= DAY_COUNT; day++) {
      if (!holidays.includes(day) && !ctx.availability[staff.name]?.[day]) {
        availableDays.push(day);
      }
    }
    
    // 15日を4等分に近づけて配置（3-4-4-4日間隔）
    const step = Math.floor(DAY_COUNT / targetHolidayCount); // 約3.75
    for (let i = 0; i < needed && i < availableDays.length; i++) {
      const idealDay = Math.floor((i + 1) * step);
      const day = availableDays.find(d => d >= idealDay) || availableDays[i];
      if (day) {
        holidays.push(day);
      }
    }
  }
  
  return holidays.sort((a, b) => a - b);
}

function assignNightForGroup_(ctx, pattern, code, candidates, label) {
  if (!candidates.length) {
    ctx.warnings.push(`${label} の Night 候補が存在しません (code=${code})`);
    return;
  }
  const assignCount = {};
  for (let day = 1; day <= DAY_COUNT; day++) {
    let loopGuard = 0;
    while (hasRemainingDemandForPattern_(ctx, day, pattern)) {
      loopGuard += 1;
      if (loopGuard > candidates.length + 5) {
        ctx.warnings.push(`${label} Night day${day} で候補不足`);
        break;
      }
      const available = candidates
        .filter(staff => canAssignPattern_(ctx, staff, day, code, { allowNightMinShortage: true }))
        .sort((a, b) => {
          const aCount = assignCount[a.name] || 0;
          const bCount = assignCount[b.name] || 0;
          if (aCount !== bCount) return aCount - bCount;
          const aHours = getAssignedHoursForStaff_(ctx, a.name);
          const bHours = getAssignedHoursForStaff_(ctx, b.name);
          if (aHours !== bHours) return aHours - bHours;
          return a.name.localeCompare(b.name, 'ja');
        });
      if (!available.length) {
        break;
      }
      const chosen = available[0];
      commitAssignment_(ctx, chosen, day, code);
      assignCount[chosen.name] = (assignCount[chosen.name] || 0) + 1;
    }
  }
}

function hasUnmetNightNeed_(ctx, pattern) {
  for (let day = 1; day <= DAY_COUNT; day++) {
    if (hasRemainingDemandForPattern_(ctx, day, pattern)) {
      return true;
    }
  }
  return false;
}

/**
 * 調理フェーズ（シンプル版：国島・山内・直美の組み合わせを順番に入れる）
 */
function phaseCook_(ctx) {
  Logger.log('🍳 Phase COOK 開始');
  const cookCodes = CONFIG.priority.cookCodes.filter(code => ctx.patterns[code]);
  Logger.log(`調理コード: ${cookCodes.join(', ')}`);
  if (!cookCodes.length) {
    ctx.warnings.push('調理パターンが見つからないため Phase COOK をスキップしました');
    return;
  }

  const cookCandidates = ctx.staff.filter(
    st => st.isFullTime && cookCodes.some(code => st.shiftFlags[code])
  );
  if (!cookCandidates.length) {
    ctx.warnings.push('調理シフトに割り当て可能な正社員が見つかりません');
    return;
  }

  const cookNames = CONFIG.priority.cookStaffNames;
  const naomi = cookCandidates.find(st => st.name === cookNames.naomi);
  const kunishima = cookCandidates.find(st => st.name === cookNames.kunishima);
  const yamauchi = cookCandidates.find(st => st.name === cookNames.yamauchi);

  if (!kunishima || !yamauchi || !naomi) {
    ctx.warnings.push('国島・山内・直美のいずれかが見つからないため、調理をスキップします');
    return;
  }

  Logger.log(`調理メンバー: 国島=${kunishima.name}, 山内=${yamauchi.name}, 直美=${naomi.name}`);

  // ステップ1: 国島・山内の公休を先に4日ずつ分散配置（NG/OFFを避ける）
  Logger.log('国島の公休を配置中...');
  distributeHolidaysForStaff_(ctx, kunishima, CONFIG.limits.defaultHolidayCount);
  Logger.log('山内の公休を配置中...');
  distributeHolidaysForStaff_(ctx, yamauchi, CONFIG.limits.defaultHolidayCount);

  // ステップ2: 1日目から15日目まで、順番に「国島＆山内 → 山内＆直美 → 直美＆国島」を埋める
  // 順番: (day-1) % 3 で決める（0=国島＆山内, 1=山内＆直美, 2=直美＆国島）
  let totalAssigned = 0;
  const pairOrder = [
    { a: kunishima, b: yamauchi, label: '国島＆山内' },
    { a: yamauchi, b: naomi, label: '山内＆直美' },
    { a: naomi, b: kunishima, label: '直美＆国島' }
  ];

  for (let day = 1; day <= DAY_COUNT; day++) {
    const need = getCookNeedForDay_(ctx, day, cookCodes);
    if (need < 2) {
      Logger.log(`day${day}: need=${need} のためスキップ`);
      continue;
    }

    // すでにこの日にCOOKが入っていたらスキップ
    const alreadyCook = [kunishima, yamauchi, naomi].some(st => {
      const code = ctx.assignments[st.name].days[day - 1];
      return code && cookCodes.includes(code);
    });
    if (alreadyCook) {
      Logger.log(`day${day}: 既にCOOKが入っているためスキップ`);
      continue;
    }

    // この日の順番を決める（基本パターン）
    // ランダム性を追加: 基本パターンに小さなランダムオフセットを追加
    const randomOffset = Math.floor(Math.random() * 3); // 0, 1, 2のいずれか
    const basePatternIdx = ((day - 1) + randomOffset) % 3;
    
    // 3つのペアを順番に試す（NG/OFF/公休が入っている場合は次のペアを試す）
    let assigned = false;
    for (let offset = 0; offset < pairOrder.length; offset++) {
      const patternIdx = (basePatternIdx + offset) % pairOrder.length;
      const pair = pairOrder[patternIdx];
      const staffA = pair.a;
      const staffB = pair.b;

      // この日のスタッフA/BがNG/OFF/公休で埋まっているかチェック
      const codeA = ctx.assignments[staffA.name].days[day - 1];
      const codeB = ctx.assignments[staffB.name].days[day - 1];
      const availA = ctx.availability[staffA.name]?.[day];
      const availB = ctx.availability[staffB.name]?.[day];
      
      // NG/OFF/公休が入っている場合は次のペアを試す
      if (codeA === CONFIG.literals.ng || codeA === CONFIG.literals.off || codeA === CONFIG.literals.holiday ||
          codeB === CONFIG.literals.ng || codeB === CONFIG.literals.off || codeB === CONFIG.literals.holiday ||
          availA === CONFIG.literals.ng || availA === CONFIG.literals.off || availA === CONFIG.literals.holiday ||
          availB === CONFIG.literals.ng || availB === CONFIG.literals.off || availB === CONFIG.literals.holiday) {
        continue; // 次のペアを試す
      }

      // 2人とも同じFCコードに入れる
      for (const code of cookCodes) {
        if (!canAssignPattern_(ctx, staffA, day, code) || !canAssignPattern_(ctx, staffB, day, code)) {
          continue;
        }
        const pat = ctx.patterns[code];
        if (!hasRemainingDemandForPattern_(ctx, day, pat)) {
          continue;
        }
        commitAssignment_(ctx, staffA, day, code);
        commitAssignment_(ctx, staffB, day, code);
        Logger.log(`day${day}: ${pair.label} (${staffA.name} & ${staffB.name}) に ${code} を割り当て`);
        Logger.log(`  → ${staffA.name} の ${day}日目: ${ctx.assignments[staffA.name].days[day - 1]}`);
        Logger.log(`  → ${staffB.name} の ${day}日目: ${ctx.assignments[staffB.name].days[day - 1]}`);
        totalAssigned += 2;
        assigned = true;
        break;
      }

      if (assigned) break; // 割り当て成功したら終了
    }

    if (!assigned) {
      Logger.log(`day${day}: どのペアも割り当てできませんでした`);
    }
  }
  Logger.log(`合計 ${totalAssigned} 件の調理シフトを割り当てました`);

  // 3) 国島・山内の総労働時間が Min に満たなくなるまで、
  //    「直美が入っているCOOKシフトの日」を国島＆山内ペアに差し替え（直美は空欄のまま）
  //    両方の総労働時間が満たされた時点で、直美が残っている日数が正しい結果
  let maxIterations = 100; // 無限ループ防止
  let iteration = 0;
  
  while (iteration < maxIterations) {
    iteration += 1;
    recalcMetrics_(ctx);
    
    const kunishimaMetric = ctx.metrics[kunishima.name];
    const yamauchiMetric = ctx.metrics[yamauchi.name];
    const kunishimaShortage = Math.max(0, kunishima.minHours - kunishimaMetric.totalHours);
    const yamauchiShortage = Math.max(0, yamauchi.minHours - yamauchiMetric.totalHours);
    
    Logger.log(`試行 ${iteration}: 国島不足=${kunishimaShortage}時間, 山内不足=${yamauchiShortage}時間`);
    
    if (kunishimaShortage <= 0 && yamauchiShortage <= 0) {
      Logger.log('✅ 国島・山内の総労働時間が満たされました');
      break; // 両方とも満たされた
    }
    
    // 直美が入っているCOOKシフトの日を探して、国島＆山内ペアに置き換える
    let replaced = false;
    for (let day = 1; day <= DAY_COUNT; day++) {
      const nCode = ctx.assignments[naomi.name].days[day - 1];
      if (!nCode || !cookCodes.includes(nCode)) continue;

      const kCode = ctx.assignments[kunishima.name].days[day - 1];
      const yCode = ctx.assignments[yamauchi.name].days[day - 1];

      // すでに国島＆山内ペアならスキップ
      const kIsCook = kCode && cookCodes.includes(kCode);
      const yIsCook = yCode && cookCodes.includes(yCode);
      if (kIsCook && yIsCook) {
        continue; // 既に国島＆山内ペア
      }

      // 直美が入っている日で、国島または山内のどちらかが既に入っている場合
      // その日を国島＆山内ペアに置き換える
      const code = nCode; // 直美が使っていたコードを使う
      
      // 国島が既に入っている場合は、山内を追加
      if (kIsCook && !yIsCook) {
        if (canAssignPattern_(ctx, yamauchi, day, code) && hasRemainingDemandForPattern_(ctx, day, ctx.patterns[code])) {
          ctx.assignments[naomi.name].days[day - 1] = '';
          commitAssignment_(ctx, yamauchi, day, code);
          Logger.log(`day${day}: 直美の ${code} を山内に差し替え（国島は既に入っている）`);
          replaced = true;
          break;
        }
      }
      // 山内が既に入っている場合は、国島を追加
      else if (yIsCook && !kIsCook) {
        if (canAssignPattern_(ctx, kunishima, day, code) && hasRemainingDemandForPattern_(ctx, day, ctx.patterns[code])) {
          ctx.assignments[naomi.name].days[day - 1] = '';
          commitAssignment_(ctx, kunishima, day, code);
          Logger.log(`day${day}: 直美の ${code} を国島に差し替え（山内は既に入っている）`);
          replaced = true;
          break;
        }
      }
      // どちらも入っていない場合は、国島＆山内の両方を入れる
      else if (!kIsCook && !yIsCook) {
        if (canAssignPattern_(ctx, kunishima, day, code) && canAssignPattern_(ctx, yamauchi, day, code) && 
            hasRemainingDemandForPattern_(ctx, day, ctx.patterns[code])) {
          ctx.assignments[naomi.name].days[day - 1] = '';
          commitAssignment_(ctx, kunishima, day, code);
          commitAssignment_(ctx, yamauchi, day, code);
          Logger.log(`day${day}: 直美の ${code} を国島＆山内ペアに差し替え`);
          replaced = true;
          break;
        }
      }
    }
    
    if (!replaced) {
      Logger.log('⚠️ 直美を減らせる日がなくなりました');
      // デバッグ: 直美が入っている日を確認
      const naomiDays = [];
      for (let day = 1; day <= DAY_COUNT; day++) {
        const code = ctx.assignments[naomi.name].days[day - 1];
        if (code && cookCodes.includes(code)) {
          naomiDays.push(`${day}日(${code})`);
        }
      }
      Logger.log(`直美が残っている日: ${naomiDays.length}日 - ${naomiDays.join(', ')}`);
      break; // もう置き換えられる日がない
    }
  }
  
  if (iteration >= maxIterations) {
    Logger.log('⚠️ 最大試行回数に達しました');
  }
  
  // 最終的な総労働時間と直美の残り日数を確認
  recalcMetrics_(ctx);
  const finalKunishimaShortage = Math.max(0, kunishima.minHours - ctx.metrics[kunishima.name].totalHours);
  const finalYamauchiShortage = Math.max(0, yamauchi.minHours - ctx.metrics[yamauchi.name].totalHours);
  const naomiRemainingDays = ctx.assignments[naomi.name].days.filter(code => code && cookCodes.includes(code)).length;
  
  Logger.log(`最終結果: 国島不足=${finalKunishimaShortage}時間, 山内不足=${finalYamauchiShortage}時間, 直美残り=${naomiRemainingDays}日`);
  
  if (finalKunishimaShortage > 0) {
    ctx.warnings.push(`国島の総労働時間が ${finalKunishimaShortage} 時間不足しています`);
  }
  if (finalYamauchiShortage > 0) {
    ctx.warnings.push(`山内の総労働時間が ${finalYamauchiShortage} 時間不足しています`);
  }
}

/**
 * スタッフの公休をずらして分散配置する
 */
function distributeHolidaysForStaff_(ctx, staff, targetHolidayCount) {
  // 既にNG/OFF/公休が入っている日を確認
  const fixedDays = [];
  for (let day = 1; day <= DAY_COUNT; day++) {
    const code = ctx.assignments[staff.name].days[day - 1];
    const avail = ctx.availability[staff.name]?.[day];
    if (code === CONFIG.literals.ng || code === CONFIG.literals.off || code === CONFIG.literals.holiday ||
        avail === CONFIG.literals.ng || avail === CONFIG.literals.off || avail === CONFIG.literals.holiday) {
      fixedDays.push(day);
    }
  }
  
  // 現在の公休数を確認
  recalcMetrics_(ctx);
  const currentHolidayCount = ctx.metrics[staff.name].holidayCount;
  const daysToAdd = targetHolidayCount - currentHolidayCount;
  
  if (daysToAdd <= 0) {
    return; // 既に公休数が十分
  }
  
  // 空欄の日を探す
  const availableDays = [];
  for (let day = 1; day <= DAY_COUNT; day++) {
    if (fixedDays.includes(day)) continue;
    const code = ctx.assignments[staff.name].days[day - 1];
    if (!code || code === '') {
      const avail = ctx.availability[staff.name]?.[day];
      if (!avail || avail === '') {
        availableDays.push(day);
      }
    }
  }
  
  // 均等に分散配置（3-4日に1回程度）
  // ランダム性を追加して、毎回異なる結果を生成する
  if (availableDays.length > 0) {
    // 利用可能な日をシャッフル（ランダム性を追加）
    const shuffled = [...availableDays];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    
    // 均等に分散配置（シャッフル後の配列から選択）
    const step = Math.max(1, Math.floor(shuffled.length / targetHolidayCount));
    Logger.log(`${staff.name} の公休配置: ${daysToAdd}日を追加、利用可能日数=${availableDays.length}, step=${step}`);
    for (let i = 0; i < daysToAdd && i < shuffled.length; i++) {
      const day = shuffled[i * step] || shuffled[i];
      if (day) {
        ctx.assignments[staff.name].days[day - 1] = CONFIG.literals.holiday;
        Logger.log(`${staff.name} の ${day}日目に公休を配置`);
      }
    }
  } else {
    Logger.log(`${staff.name} の公休配置: 利用可能な日がありません`);
  }
}

/**
 * 調理スタッフ（国島・山内）の公休数と総労働時間を完全一致させる（未使用）
 */
function adjustCookStaffConstraints_(ctx, kunishima, yamauchi, cookCodes) {
  [kunishima, yamauchi].filter(Boolean).forEach(staff => {
    if (!staff) return;
    
    // 最大100回試行して制約を満たす
    for (let attempt = 0; attempt < 100; attempt++) {
      recalcMetrics_(ctx);
      const metric = ctx.metrics[staff.name];
      const targetHolidayCount = CONFIG.limits.defaultHolidayCount;
      const currentHolidayCount = metric.holidayCount;
      const currentHours = metric.totalHours;
      
      // 制約を満たしているかチェック
      const holidayOk = currentHolidayCount === targetHolidayCount;
      const hoursOk = currentHours >= staff.minHours && currentHours <= staff.maxHours;
      
      if (holidayOk && hoursOk) {
        break; // 制約を満たした
      }
      
      // 公休数を調整
      if (!holidayOk) {
        if (currentHolidayCount < targetHolidayCount) {
          // 公休を追加（空欄の日を公休に）
          const daysToAdd = targetHolidayCount - currentHolidayCount;
          const availableDays = [];
          for (let day = 1; day <= DAY_COUNT; day++) {
            const code = ctx.assignments[staff.name].days[day - 1];
            if (!code || code === '') {
              const avail = ctx.availability[staff.name]?.[day];
              if (!avail || avail === '') {
                availableDays.push(day);
              }
            }
          }
          // 均等に分散配置
          if (availableDays.length > 0) {
            const step = Math.max(1, Math.floor(availableDays.length / targetHolidayCount));
            for (let i = 0; i < daysToAdd && i < availableDays.length; i++) {
              const day = availableDays[i * step] || availableDays[i];
              if (day) {
                ctx.assignments[staff.name].days[day - 1] = CONFIG.literals.holiday;
              }
            }
          }
        } else if (currentHolidayCount > targetHolidayCount) {
          // 公休を減らす（公休の日を調理に変更）
          const daysToRemove = currentHolidayCount - targetHolidayCount;
          const holidayDays = [];
          for (let day = 1; day <= DAY_COUNT; day++) {
            const code = ctx.assignments[staff.name].days[day - 1];
            if (code === CONFIG.literals.holiday || code === CONFIG.literals.ng) {
              holidayDays.push(day);
            }
          }
          // 調理のneedがある日を優先して公休を減らす
          for (let i = 0; i < daysToRemove && i < holidayDays.length; i++) {
            const day = holidayDays[i];
            const need = getCookNeedForDay_(ctx, day, cookCodes);
            if (need > 0) {
              const code = cookCodes.find(c => canAssignPattern_(ctx, staff, day, c));
              if (code && hasRemainingDemandForPattern_(ctx, day, ctx.patterns[code])) {
                ctx.assignments[staff.name].days[day - 1] = code;
              }
            }
          }
        }
      }
      
      // 総労働時間を調整（公休数を維持しながら）
      recalcMetrics_(ctx);
      const newMetric = ctx.metrics[staff.name];
      const newHours = newMetric.totalHours;
      
      if (!hoursOk) {
        if (newHours < staff.minHours) {
          // 総労働時間が少なすぎる場合、調理を追加
          const hoursToAdd = staff.minHours - newHours;
          const availableDays = [];
          for (let day = 1; day <= DAY_COUNT; day++) {
            const code = ctx.assignments[staff.name].days[day - 1];
            if (code === CONFIG.literals.holiday || code === CONFIG.literals.ng || !code) {
              const need = getCookNeedForDay_(ctx, day, cookCodes);
              if (need > 0) {
                const cookCode = cookCodes.find(c => canAssignPattern_(ctx, staff, day, c));
                if (cookCode && hasRemainingDemandForPattern_(ctx, day, ctx.patterns[cookCode])) {
                  const pattern = ctx.patterns[cookCode];
                  const hours = getPatternHourCount_(pattern);
                  availableDays.push({ day, code: cookCode, hours });
                }
              }
            }
          }
          // 労働時間が足りる分だけ追加（公休数を維持）
          let addedHours = 0;
          for (const item of availableDays) {
            if (addedHours >= hoursToAdd) break;
            recalcMetrics_(ctx);
            const tempMetric = ctx.metrics[staff.name];
            if (tempMetric.holidayCount >= targetHolidayCount) {
              // 公休が多すぎる場合は減らす
              ctx.assignments[staff.name].days[item.day - 1] = item.code;
              addedHours += item.hours;
            }
          }
        } else if (newHours > staff.maxHours) {
          // 総労働時間が多すぎる場合、調理を減らす
          const hoursToRemove = newHours - staff.maxHours;
          const workDays = [];
          for (let day = 1; day <= DAY_COUNT; day++) {
            const code = ctx.assignments[staff.name].days[day - 1];
            if (code && cookCodes.includes(code)) {
              const pattern = ctx.patterns[code];
              const hours = getPatternHourCount_(pattern);
              workDays.push({ day, hours });
            }
          }
          // 労働時間が減る分だけ公休に変更（公休数を維持）
          let removedHours = 0;
          for (const item of workDays) {
            if (removedHours >= hoursToRemove) break;
            recalcMetrics_(ctx);
            const tempMetric = ctx.metrics[staff.name];
            if (tempMetric.holidayCount <= targetHolidayCount) {
              // 公休が少なすぎる場合は増やす
              ctx.assignments[staff.name].days[item.day - 1] = CONFIG.literals.holiday;
              removedHours += item.hours;
            }
          }
        }
      }
    }
  });
}

/**
 * 調理の組み合わせルールに基づいて候補を選ぶ
 */
function findCookCombination_(ctx, day, candidates, cookCodes, alreadyAssigned, specialists, isWeekday) {
  const result = [];
  const need = getCookNeedForDay_(ctx, day, cookCodes);
  const currentCount = alreadyAssigned.length;
  const remaining = need - currentCount;

  if (remaining <= 0) {
    return [];
  }

  // 組み合わせルールを優先順位で試す
  for (const combo of CONFIG.priority.cookCombinations) {
    if (combo.weekdayOnly && !isWeekday) {
      continue;
    }
    const comboMembers = combo.members.filter(name => candidates.some(st => st.name === name));
    if (comboMembers.length < 2) {
      continue;
    }

    // 既に割り当て済みのメンバーを確認
    const assignedInCombo = comboMembers.filter(name => alreadyAssigned.includes(name));
    const unassignedInCombo = comboMembers.filter(name => !alreadyAssigned.includes(name));

    // コンビのメンバーが既に割り当て済みなら、残りを追加
    if (assignedInCombo.length > 0 && unassignedInCombo.length > 0 && remaining > 0) {
      const toAdd = unassignedInCombo.slice(0, remaining);
      for (const name of toAdd) {
        const staff = candidates.find(st => st.name === name);
        if (staff) {
          const code = cookCodes.find(code => canAssignPattern_(ctx, staff, day, code));
          if (code) {
            result.push({ staff, code });
          }
        }
      }
      if (result.length > 0) {
        return result;
      }
    }
  }

  // 組み合わせルールに該当しない場合は、専任を優先して選ぶ
  const available = candidates
    .filter(st => !alreadyAssigned.includes(st.name))
    .map(staff => {
      const code = cookCodes.find(code => canAssignPattern_(ctx, staff, day, code));
      return code ? { staff, code } : null;
    })
    .filter(Boolean)
    .sort((a, b) => {
      const aSpec = specialists.has(a.staff.name);
      const bSpec = specialists.has(b.staff.name);
      if (aSpec !== bSpec) return aSpec ? -1 : 1;
      return a.staff.name.localeCompare(b.staff.name, 'ja');
    });

  return available.slice(0, remaining);
}

/**
 * 調理の need を取得
 */
function getCookNeedForDay_(ctx, day, cookCodes) {
  let maxNeed = 0;
  cookCodes.forEach(code => {
    const pattern = ctx.patterns[code];
    if (!pattern) return;
    Object.keys(pattern.cover).forEach(department => {
      if (pattern.cover[department] && pattern.cover[department].length > 0) {
        const hour = pattern.cover[department][0];
        const need = ctx.demand[day]?.[department]?.[hour] || 0;
        maxNeed = Math.max(maxNeed, need);
      }
    });
  });
  return maxNeed;
}

/**
 * 平日かどうかを判定（簡易版: 土日を避ける）
 */
function isWeekday_(day) {
  // 簡易実装: 実際の日付に基づく判定が必要な場合は改善
  // ここでは 1-15 日のうち、3, 6, 9, 12, 15 日目を週末と仮定（要調整）
  const weekendDays = [3, 6, 9, 12, 15];
  return !weekendDays.includes(day);
}

/**
 * パターンの労働時間を取得（ShiftPattern の「労働時間」列から）
 */
function getPatternHourCount_(pattern) {
  // ShiftPattern の「労働時間」列から取得
  const raw = pattern.raw || [];
  const hoursColIndex = pattern.hoursColIndex !== undefined ? pattern.hoursColIndex : 10; // デフォルト: L列（0-indexedで10）
  
  if (raw.length > hoursColIndex) {
    const hours = Number(raw[hoursColIndex]) || 0;
    if (hours > 0) {
      return hours;
    }
  }
  
  // フォールバック: デフォルト値（本来は避けるべき）
  Logger.log(`警告: シフトコード ${pattern.code} の労働時間が取得できません。hoursColIndex=${hoursColIndex}, raw.length=${raw.length}`);
  return 8; // デフォルト値（本来は避けるべき）
}

/**
 * バー+夕食フェーズ（FDB-S + 夕食帯を同時配置）
 */
function phaseBarAndDinner_(ctx) {
  const barCodes = CONFIG.priority.barCodes.filter(code => ctx.patterns[code]);
  const dinnerCodes = Object.keys(ctx.patterns).filter(code => {
    const pattern = ctx.patterns[code];
    return pattern.cover && pattern.cover['夕食'];
  });
  
  const candidates = ctx.staff.filter(st => st.isFullTime && 
    (barCodes.some(code => st.shiftFlags[code]) || 
     dinnerCodes.some(code => st.shiftFlags[code])));
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    // バーのneedを確認
    const barNeed = getDemandForDepartmentAndTime_(ctx, day, 'バー', 20, 23);
    const dinnerNeed = getDemandForDepartmentAndTime_(ctx, day, '夕食', 18, 20);
    
    // バーがある日はFDB-Sを優先
    if (barNeed > 0) {
      const available = candidates
        .filter(st => {
          const current = ctx.assignments[st.name].days[day - 1];
          return !current || current === '';
        })
        .filter(st => barCodes.some(code => canAssignPattern_(ctx, st, day, code)))
        .sort((a, b) => {
          const aHours = getAssignedHoursForStaff_(ctx, a.name);
          const bHours = getAssignedHoursForStaff_(ctx, b.name);
          return aHours - bHours;
        });
      
      for (let i = 0; i < barNeed && i < available.length; i++) {
        const staff = available[i];
        const code = barCodes.find(c => canAssignPattern_(ctx, staff, day, c));
        if (code) {
          commitAssignment_(ctx, staff, day, code);
        }
      }
    }
    
    // 夕食のneedを確認（バーと重複しないように）
    if (dinnerNeed > 0) {
      let assigned = 0;
      candidates.forEach(staff => {
        const code = ctx.assignments[staff.name].days[day - 1];
        if (code && dinnerCodes.includes(code)) {
          assigned += 1;
        }
      });
      
      const remaining = dinnerNeed - assigned;
      if (remaining > 0) {
        const available = candidates
          .filter(st => {
            const current = ctx.assignments[st.name].days[day - 1];
            return !current || current === '';
          })
          .filter(st => dinnerCodes.some(code => canAssignPattern_(ctx, st, day, code)))
          .sort((a, b) => {
            const aHours = getAssignedHoursForStaff_(ctx, a.name);
            const bHours = getAssignedHoursForStaff_(ctx, b.name);
            return aHours - bHours;
          });
        
        for (let i = 0; i < remaining && i < available.length; i++) {
          const staff = available[i];
          const code = dinnerCodes.find(c => canAssignPattern_(ctx, staff, day, c));
          if (code) {
            commitAssignment_(ctx, staff, day, code);
          }
        }
      }
    }
  }
}

/**
 * 朝食コアフェーズ（6-9の朝食帯のみを正社員で配置）
 */
function phaseBreakfastCore_(ctx) {
  const codes = CONFIG.priority.breakfastCoreCodes.filter(code => ctx.patterns[code]);
  if (!codes.length) return;
  
  const candidates = ctx.staff.filter(st => st.isFullTime && codes.some(code => st.shiftFlags[code]));
  if (!candidates.length) return;
  
  // 6-9の朝食帯のDemandを満たす（各時間帯ごとにチェック）
  for (let day = 1; day <= DAY_COUNT; day++) {
    // 6-9時の範囲でneedがあるかチェック（6-9時の範囲でneed>0の時間帯がある場合のみ実行）
    let hasNeedInRange = false;
    for (let hour = 6; hour <= 9; hour++) {
      const need = ctx.demand[day]?.['朝食']?.[hour] || 0;
      if (need > 0) {
        hasNeedInRange = true;
        break;
      }
    }
    if (!hasNeedInRange) continue; // 6-9時にneedがない場合はスキップ
    
    // 6-9時の各時間帯で最大のneedを取得
    let maxNeed = 0;
    for (let hour = 6; hour <= 9; hour++) {
      const need = ctx.demand[day]?.['朝食']?.[hour] || 0;
      maxNeed = Math.max(maxNeed, need);
    }
    if (maxNeed <= 0) continue;
    
    // 既に割り当てられている人数を確認
    let assigned = 0;
    candidates.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && codes.includes(code)) {
        assigned += 1;
      }
    });
    
    let remaining = maxNeed - assigned;
    if (remaining <= 0) continue;
    
    // 候補を選んで割り当て
    const available = candidates
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => codes.some(code => canAssignPattern_(ctx, st, day, code)))
      .sort((a, b) => {
        const aHours = getAssignedHoursForStaff_(ctx, a.name);
        const bHours = getAssignedHoursForStaff_(ctx, b.name);
        return aHours - bHours;
      });
    
    for (let i = 0; i < remaining && i < available.length; i++) {
      const staff = available[i];
      const code = codes.find(c => {
        if (!canAssignPattern_(ctx, staff, day, c)) return false;
        // 割り当て前にneedを再チェック（各時間帯ごとに）
        const pattern = ctx.patterns[c];
        if (!pattern.cover || !pattern.cover['朝食']) return false;
        // このシフトがカバーする時間帯で、すべてneedがあるかチェック
        const coveredHours = pattern.cover['朝食'];
        for (const hour of coveredHours) {
          if (hour >= 6 && hour <= 9) {
            const need = ctx.demand[day]?.['朝食']?.[hour] || 0;
            const have = calcHave_(ctx, day, '朝食', hour);
            if (have >= need) return false; // この時間帯は既に満たされている
          }
        }
        return true;
      });
      if (code) {
        commitAssignment_(ctx, staff, day, code);
        // 割り当て後にneedを再計算
        const newAssigned = candidates.filter(st => {
          const c = ctx.assignments[st.name].days[day - 1];
          return c && codes.includes(c);
        }).length;
        if (newAssigned >= maxNeed) break; // needを満たしたら終了
        remaining = maxNeed - newAssigned;
      }
    }
  }
}

/**
 * ロビー午前フェーズ（ロング→ショート交互、村松/三本木を中心に配置）
 */
function phaseLobbyAm_(ctx) {
  const coreNames = CONFIG.priority.lobbyAmCore;
  const coreStaff = ctx.staff.filter(st => coreNames.includes(st.name) && st.isFullTime);
  const otherStaff = ctx.staff.filter(st => !coreNames.includes(st.name) && st.isFullTime);
  
  // ロビーAMのシフトコードを取得（ShiftPatternから）
  const lobbyAmCodes = Object.keys(ctx.patterns).filter(code => {
    const pattern = ctx.patterns[code];
    return pattern.cover && pattern.cover['ロビー'] && 
           pattern.cover['ロビー'].some(h => h >= 8 && h <= 13);
  });
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    const need = getDemandForDepartmentAndTime_(ctx, day, 'ロビー', 8, 13);
    if (need <= 0) continue;
    
    // 既に割り当てられている人数を確認
    let assigned = 0;
    ctx.staff.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && lobbyAmCodes.includes(code)) {
        assigned += 1;
      }
    });
    
    let remaining = need - assigned;
    if (remaining <= 0) continue;
    
    // コアスタッフを優先して交互に配置
    const availableCore = coreStaff
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => lobbyAmCodes.some(code => canAssignPattern_(ctx, st, day, code)));
    
    // 交互に配置（簡易実装）
    for (let i = 0; i < remaining && i < availableCore.length; i++) {
      const staff = availableCore[i % availableCore.length];
      const code = lobbyAmCodes.find(c => canAssignPattern_(ctx, staff, day, c));
      if (code) {
        commitAssignment_(ctx, staff, day, code);
        remaining -= 1;
      }
    }
    
    // 残りは他のスタッフで埋める
    if (remaining > 0) {
      const available = otherStaff
        .filter(st => {
          const current = ctx.assignments[st.name].days[day - 1];
          return !current || current === '';
        })
        .filter(st => lobbyAmCodes.some(code => canAssignPattern_(ctx, st, day, code)));
      
      for (let i = 0; i < remaining && i < available.length; i++) {
        const staff = available[i];
        const code = lobbyAmCodes.find(c => canAssignPattern_(ctx, staff, day, c));
        if (code) {
          commitAssignment_(ctx, staff, day, code);
        }
      }
    }
  }
}

/**
 * ロビー午後フェーズ
 */
function phaseLobbyPm_(ctx) {
  const lobbyPmCodes = Object.keys(ctx.patterns).filter(code => {
    const pattern = ctx.patterns[code];
    return pattern.cover && pattern.cover['ロビー'] && 
           pattern.cover['ロビー'].some(h => h >= 13 && h <= 22);
  });
  
  const candidates = ctx.staff.filter(st => st.isFullTime && 
    lobbyPmCodes.some(code => st.shiftFlags[code]));
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    const need = getDemandForDepartmentAndTime_(ctx, day, 'ロビー', 13, 22);
    if (need <= 0) continue;
    
    let assigned = 0;
    candidates.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && lobbyPmCodes.includes(code)) {
        assigned += 1;
      }
    });
    
    const remaining = need - assigned;
    if (remaining <= 0) continue;
    
    const available = candidates
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => lobbyPmCodes.some(code => canAssignPattern_(ctx, st, day, code)))
      .sort((a, b) => {
        const aHours = getAssignedHoursForStaff_(ctx, a.name);
        const bHours = getAssignedHoursForStaff_(ctx, b.name);
        return aHours - bHours;
      });
    
    for (let i = 0; i < remaining && i < available.length; i++) {
      const staff = available[i];
      const code = lobbyPmCodes.find(c => canAssignPattern_(ctx, staff, day, c));
      if (code) {
        commitAssignment_(ctx, staff, day, code);
      }
    }
  }
}

/**
 * 夕食フェーズ
 */
function phaseDinner_(ctx) {
  const dinnerCodes = Object.keys(ctx.patterns).filter(code => {
    const pattern = ctx.patterns[code];
    return pattern.cover && pattern.cover['夕食'];
  });
  
  const candidates = ctx.staff.filter(st => st.isFullTime && 
    dinnerCodes.some(code => st.shiftFlags[code]));
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    const need = getDemandForDepartmentAndTime_(ctx, day, '夕食', 16, 22);
    if (need <= 0) continue;
    
    let assigned = 0;
    candidates.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && dinnerCodes.includes(code)) {
        assigned += 1;
      }
    });
    
    const remaining = need - assigned;
    if (remaining <= 0) continue;
    
    const available = candidates
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => dinnerCodes.some(code => canAssignPattern_(ctx, st, day, code)))
      .sort((a, b) => {
        const aHours = getAssignedHoursForStaff_(ctx, a.name);
        const bHours = getAssignedHoursForStaff_(ctx, b.name);
        return aHours - bHours;
      });
    
    for (let i = 0; i < remaining && i < available.length; i++) {
      const staff = available[i];
      const code = dinnerCodes.find(c => canAssignPattern_(ctx, staff, day, c));
      if (code) {
        commitAssignment_(ctx, staff, day, code);
      }
    }
  }
}

/**
 * 朝食フェーズ（残りを配置）
 */
function phaseBreakfast_(ctx) {
  const codes = CONFIG.priority.breakfastCoreCodes.filter(code => ctx.patterns[code]);
  if (!codes.length) return;
  
  const candidates = ctx.staff.filter(st => st.isFullTime && codes.some(code => st.shiftFlags[code]));
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    // 5-11時の各時間帯で最大のneedを取得
    let maxNeed = 0;
    for (let hour = 5; hour <= 11; hour++) {
      const need = ctx.demand[day]?.['朝食']?.[hour] || 0;
      maxNeed = Math.max(maxNeed, need);
    }
    if (maxNeed <= 0) continue;
    
    let assigned = 0;
    candidates.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && codes.includes(code)) {
        assigned += 1;
      }
    });
    
    let remaining = maxNeed - assigned;
    if (remaining <= 0) continue;
    
    const available = candidates
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => codes.some(code => canAssignPattern_(ctx, st, day, code)))
      .sort((a, b) => {
        const aHours = getAssignedHoursForStaff_(ctx, a.name);
        const bHours = getAssignedHoursForStaff_(ctx, b.name);
        return aHours - bHours;
      });
    
    for (let i = 0; i < remaining && i < available.length; i++) {
      const staff = available[i];
      const code = codes.find(c => {
        if (!canAssignPattern_(ctx, staff, day, c)) return false;
        // 割り当て前にneedを再チェック（各時間帯ごとに）
        const pattern = ctx.patterns[c];
        if (!pattern.cover || !pattern.cover['朝食']) return false;
        // このシフトがカバーする時間帯で、すべてneedがあるかチェック
        const coveredHours = pattern.cover['朝食'];
        for (const hour of coveredHours) {
          if (hour >= 5 && hour <= 11) {
            const need = ctx.demand[day]?.['朝食']?.[hour] || 0;
            const have = calcHave_(ctx, day, '朝食', hour);
            if (have >= need) return false; // この時間帯は既に満たされている
          }
        }
        return true;
      });
      if (code) {
        commitAssignment_(ctx, staff, day, code);
        // 割り当て後にneedを再計算
        const newAssigned = candidates.filter(st => {
          const c = ctx.assignments[st.name].days[day - 1];
          return c && codes.includes(c);
        }).length;
        if (newAssigned >= maxNeed) break; // needを満たしたら終了
        remaining = maxNeed - newAssigned;
      }
    }
  }
}

/**
 * 清掃フェーズ
 */
function phaseCleaning_(ctx) {
  const cleaningCodes = Object.keys(ctx.patterns).filter(code => {
    const pattern = ctx.patterns[code];
    return pattern.cover && pattern.cover['清掃'];
  });
  
  const candidates = ctx.staff.filter(st => st.isFullTime && 
    cleaningCodes.some(code => st.shiftFlags[code]));
  
  for (let day = 1; day <= DAY_COUNT; day++) {
    const need = getDemandForDepartmentAndTime_(ctx, day, '清掃', 9, 15);
    if (need <= 0) continue;
    
    let assigned = 0;
    candidates.forEach(staff => {
      const code = ctx.assignments[staff.name].days[day - 1];
      if (code && cleaningCodes.includes(code)) {
        assigned += 1;
      }
    });
    
    const remaining = need - assigned;
    if (remaining <= 0) continue;
    
    const available = candidates
      .filter(st => {
        const current = ctx.assignments[st.name].days[day - 1];
        return !current || current === '';
      })
      .filter(st => cleaningCodes.some(code => canAssignPattern_(ctx, st, day, code)))
      .sort((a, b) => {
        const aHours = getAssignedHoursForStaff_(ctx, a.name);
        const bHours = getAssignedHoursForStaff_(ctx, b.name);
        return aHours - bHours;
      });
    
    for (let i = 0; i < remaining && i < available.length; i++) {
      const staff = available[i];
      const code = cleaningCodes.find(c => canAssignPattern_(ctx, staff, day, c));
      if (code) {
        commitAssignment_(ctx, staff, day, code);
      }
    }
  }
}

/**
 * パート・アルバイトフェーズ（残 need をゼロにする）
 */
function phasePartTimer_(ctx) {
  const partTimers = ctx.staff.filter(st => !st.isFullTime);
  if (!partTimers.length) return;
  
  // すべての部門・時間帯で残っているneedを埋める
  for (let day = 1; day <= DAY_COUNT; day++) {
    const departments = Object.keys(ctx.demand[day] || {});
    departments.forEach(dep => {
      HOURS_SEQUENCE.forEach(hour => {
        const need = ctx.demand[day][dep]?.[hour] || 0;
        const have = calcHave_(ctx, day, dep, hour);
        const remaining = need - have;
        
        if (remaining <= 0) return;
        
        // この部門・時間帯をカバーするシフトコードを探す
        const coveringCodes = Object.keys(ctx.patterns).filter(code => {
          const pattern = ctx.patterns[code];
          return pattern.cover && pattern.cover[dep] && pattern.cover[dep].includes(hour);
        });
        
        // パートタイマーで割り当て可能な候補を探す
        const available = partTimers
          .filter(st => {
            const current = ctx.assignments[st.name].days[day - 1];
            return !current || current === '';
          })
          .filter(st => coveringCodes.some(code => st.shiftFlags[code]))
          .filter(st => coveringCodes.some(code => canAssignPattern_(ctx, st, day, code)))
          .sort((a, b) => {
            const aHours = getAssignedHoursForStaff_(ctx, a.name);
            const bHours = getAssignedHoursForStaff_(ctx, b.name);
            return aHours - bHours;
          });
        
        for (let i = 0; i < remaining && i < available.length; i++) {
          const staff = available[i];
          const code = coveringCodes.find(c => canAssignPattern_(ctx, staff, day, c));
          if (code) {
            commitAssignment_(ctx, staff, day, code);
          }
        }
      });
    });
  }
}

/**
 * 部門・時間帯のDemandを取得
 */
function getDemandForDepartmentAndTime_(ctx, day, department, startHour, endHour) {
  let maxNeed = 0;
  for (let hour = startHour; hour <= endHour; hour++) {
    const need = ctx.demand[day]?.[department]?.[hour] || 0;
    maxNeed = Math.max(maxNeed, need);
  }
  return maxNeed;
}

/**
 * Shift_15a へ assignments を書き戻す
 * @param {ShiftContext} ctx
 */
function writeShiftSheet_(ctx) {
  try {
    Logger.log('📝 writeShiftSheet_ 開始');
    const actualSpreadsheetId = ctx.spreadsheet.getId();
    Logger.log(`実際のスプレッドシートID: ${actualSpreadsheetId}`);
    Logger.log(`期待されるID: ${SPREADSHEET_ID}`);
    Logger.log(`IDが一致しているか: ${actualSpreadsheetId === SPREADSHEET_ID}`);
    Logger.log(`出力シート名: ${CONFIG.sheets.output}`);
    
    if (actualSpreadsheetId !== SPREADSHEET_ID) {
      Logger.log(`⚠️ 警告: スプレッドシートIDが一致しません！`);
      Logger.log(`  実際: ${actualSpreadsheetId}`);
      Logger.log(`  期待: ${SPREADSHEET_ID}`);
    }
    
    const sheet = ctx.spreadsheet.getSheetByName(CONFIG.sheets.output);
    if (!sheet) {
      // 利用可能なシート名を確認
      const allSheets = ctx.spreadsheet.getSheets();
      const sheetNames = allSheets.map(s => s.getName());
      Logger.log(`利用可能なシート: ${sheetNames.join(', ')}`);
      throw new Error(`出力シート ${CONFIG.sheets.output} が見つかりません`);
    }
    
    Logger.log(`出力シート "${CONFIG.sheets.output}" が見つかりました`);

    const values = sheet.getDataRange().getDisplayValues();
    if (!values.length) {
      throw new Error('出力シートが空です');
    }

    const header = values[0];
    Logger.log(`ヘッダー: ${header.slice(0, 10).join(', ')}...`);
    
    const dayColumnIndices = extractDayColumns_(header);
    Logger.log(`日付列インデックス: ${dayColumnIndices.join(', ')}`);
    
    if (dayColumnIndices.length !== DAY_COUNT) {
      Logger.log(`⚠️ 日付列数が不正です: ${dayColumnIndices.length} (期待値: ${DAY_COUNT})`);
    }
    
    const totalCol = header.indexOf('総労働時間');
    const workdayCol = header.indexOf('出勤日数');
    const holidayCol = header.indexOf('休日数');
    
    Logger.log(`日付列数: ${dayColumnIndices.length}, 総労働時間列: ${totalCol}, 出勤日数列: ${workdayCol}, 休日数列: ${holidayCol}`);

  // 1) シフト列 + Q/R/S 列の背景色をまとめてリセット（前回の色が残らないようにする）
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 1) {
    // 2行目以降・B列以降をすべてクリア（罫線や値はそのまま）
    sheet.getRange(2, 2, lastRow - 1, lastCol - 1).setBackground(null);
  }

    Logger.log(`書き込み対象行数: ${values.length - 1}`);
    Logger.log(`assignments のキー数: ${Object.keys(ctx.assignments).length}`);
    Logger.log(`assignments のキー例: ${Object.keys(ctx.assignments).slice(0, 5).join(', ')}`);
    
    let writtenCount = 0;
    let skippedCount = 0;
    
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const name = row[0];
      if (!name) {
        skippedCount++;
        continue;
      }
      
      const assignment = ctx.assignments[name];
      if (!assignment) {
        skippedCount++;
        if (name === '国島' || name === '直美' || name === '山内') {
          Logger.log(`⚠️ ${name} の assignment が見つかりません`);
        }
        continue;
      }
      const staff = ctx.staff.find(st => st.name === name);
      if (!staff) {
        continue;
      }

      // 2) シフトコードを書き込み（色付け付き）
      if (name === '国島' || name === '直美' || name === '山内') {
        const codes = assignment.days.map((code, idx) => `${idx + 1}日=${code || '(空)'}`).join(', ');
        Logger.log(`📝 ${name} を書き込み中: ${codes}`);
      }
      
      dayColumnIndices.forEach((col, idx) => {
        if (idx >= assignment.days.length) return;
        
        const code = assignment.days[idx] || '';
        const cellRow = r + 1;
        const cellCol = col + 1;
        const cell = sheet.getRange(cellRow, cellCol);
        
        // 値を書き込み
        cell.setValue(code);
        writtenCount++;
        
        // デバッグ: 国島・直美・山内の書き込みを確認
        if ((name === '国島' || name === '直美' || name === '山内') && code) {
          Logger.log(`  → ${name} の ${idx + 1}日目 (行${cellRow}, 列${cellCol}) に "${code}" を書き込み`);
        }
      
      // 空欄の場合は背景色をクリア
      if (!code || code === '') {
        cell.setBackground(CONFIG.colors.cleared);
        return;
      }
      
      // 公休/OFF/NG の色付け
      if (code === CONFIG.literals.holiday) {
        cell.setBackground(CONFIG.colors.holiday);
      } else if (code === CONFIG.literals.off) {
        cell.setBackground(CONFIG.colors.off);
      } else if (code === CONFIG.literals.ng) {
        cell.setBackground(CONFIG.colors.ng);
      } else {
        // ナイトと調理の色付け（シフトコードがある場合のみ）
        if (CONFIG.priority.nightCodes.includes(code)) {
          cell.setBackground(CONFIG.colors.night);
        } else if (CONFIG.priority.cookCodes.includes(code)) {
          cell.setBackground(CONFIG.colors.cook);
        } else {
          // その他のシフトコードは背景色をクリア
          cell.setBackground(CONFIG.colors.cleared);
        }
      }
    });

    // Q/R/S 列（総労働時間/出勤日数/休日数）を書き込み
    const metric = ctx.metrics[name];
    if (totalCol >= 0) {
      sheet.getRange(r + 1, totalCol + 1).setValue(metric.totalHours);
      // 色分け: Min/Max の間は正常、超過は赤、未達は黄色
      const range = sheet.getRange(r + 1, totalCol + 1);
      if (metric.totalHours > staff.maxHours) {
        range.setBackground(CONFIG.colors.overMax);
      } else if (metric.totalHours < staff.minHours) {
        range.setBackground(CONFIG.colors.underMin);
      } else {
        range.setBackground(CONFIG.colors.normal);
      }
    }
    if (workdayCol >= 0) {
      sheet.getRange(r + 1, workdayCol + 1).setValue(metric.workDays);
      // 出勤日数は Min/Max の範囲外なら警告色（簡易実装）
    }
    if (holidayCol >= 0) {
      sheet.getRange(r + 1, holidayCol + 1).setValue(metric.holidayCount);
      // 正社員の場合、公休4日を期待
      if (staff.isFullTime) {
        const range = sheet.getRange(r + 1, holidayCol + 1);
        if (metric.holidayCount !== CONFIG.limits.defaultHolidayCount) {
          range.setBackground(CONFIG.colors.warning);
        } else {
          range.setBackground(CONFIG.colors.normal);
        }
      }
    }
    
    Logger.log(`✅ 書き込み完了: ${writtenCount} セル, スキップ: ${skippedCount} 行`);
  } catch (error) {
    Logger.log(`❌ writeShiftSheet_ エラー: ${error.toString()}`);
    Logger.log(`スタック: ${error.stack}`);
    throw error;
  }
}

/**
 * Line_* シートへ need / have を書き出す
 */
function writeLineSheets_(ctx) {
  for (let day = 1; day <= DAY_COUNT; day++) {
    const sheetName = `${CONFIG.sheets.linePrefix}${day}`;
    const sheet = ctx.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      continue;
    }
    sheet.clear();
    writeLineSheetForDay_(sheet, ctx, day);
  }
}

/**
 * need/have を時間別に展開して書き込む
 */
function writeLineSheetForDay_(sheet, ctx, day) {
  const demand = ctx.demand[day];
  const header = ['Hour'];
  const departments = Object.keys(demand);
  departments.forEach(dep => {
    header.push(`need_${dep}`);
    header.push(`have_${dep}`);
  });
  const rows = [header];

  HOURS_SEQUENCE.forEach(hour => {
    const row = [`${hour}:00`];
    departments.forEach(dep => {
      const needVal = demand[dep]?.[hour] || 0;
      const haveVal = calcHave_(ctx, day, dep, hour);
      row.push(needVal);
      row.push(haveVal);
    });
    rows.push(row);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  const needColumns = [];
  const haveColumns = [];
  header.forEach((title, idx) => {
    if (title.startsWith('need_')) needColumns.push(idx + 1);
    if (title.startsWith('have_')) haveColumns.push(idx + 1);
  });
  needColumns.forEach(col => sheet.getRange(2, col, rows.length - 1, 1).setBackground(CONFIG.colors.needColumn));
  haveColumns.forEach(col => sheet.getRange(2, col, rows.length - 1, 1).setBackground(CONFIG.colors.haveColumn));

  highlightNeedGap_(sheet, rows);
}

/**
 * need ≠ have を黄色で塗る
 */
function highlightNeedGap_(sheet, rows) {
  for (let r = 2; r <= rows.length; r++) {
    const row = rows[r - 1];
    for (let c = 2; c <= row.length; c += 2) {
      const need = Number(row[c - 1]);
      const have = Number(row[c]);
      if (need !== have) {
        sheet.getRange(r, c - 1, 1, 2).setBackground(CONFIG.colors.warning);
      }
    }
  }
}

/**
 * assignments から have を計算
 */
function calcHave_(ctx, day, department, hour) {
  let count = 0;
  Object.keys(ctx.assignments).forEach(name => {
    const code = ctx.assignments[name].days[day - 1];
    if (!code) {
      return;
    }
    const pattern = ctx.patterns[code];
    if (!pattern || !pattern.cover[department]) {
      return;
    }
    if (pattern.cover[department].includes(hour)) {
      count += 1;
    }
  });
  return count;
}

/**
 * StaffDB 読み込み
 */
function readStaffDb_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.sheets.staffDB);
  if (!sheet) {
    throw new Error(`StaffDB シート ${CONFIG.sheets.staffDB} が見つかりません`);
  }
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return [];
  }
  const header = values[0];
  const colIndex = createHeaderIndex_(header);
  const result = [];
  const knownHeaders = new Set(Object.values(CONFIG.colHints).flat());

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[colIndex.name] || row[colIndex.staffId] || `row_${r}`;
    const division = row[colIndex.division] || '';
    const minHours = parseNumber_(row[colIndex.min]);
    const maxHours = parseNumber_(row[colIndex.max]);
    const isFullTime = normalizeBoolean_(row[colIndex.fulltimeFlag]);

    const shiftFlags = {};
    header.forEach((title, idx) => {
      if (idx <= colIndex.max || knownHeaders.has(title)) {
        return;
      }
      shiftFlags[title] = normalizeBoolean_(row[idx]);
    });

    result.push({
      id: row[colIndex.staffId] || `auto_${r}`,
      name,
      division,
      isFullTime,
      minHours: isNaN(minHours) ? 0 : minHours,
      maxHours: isNaN(maxHours) ? 0 : maxHours,
      shiftFlags,
      isPartTimer: division.indexOf('パート') >= 0 || division.indexOf('アルバイト') >= 0
    });
  }
  return result;
}

/**
 * Availability 読み込み
 */
function readAvailability_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.sheets.availability);
  if (!sheet) {
    throw new Error(`Availability シート ${CONFIG.sheets.availability} が見つかりません`);
  }
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {};
  }
  const header = values[0];
  const dayCols = extractDayColumns_(header);
  const availability = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[0];
    if (!name) continue;
    availability[name] = {};
    dayCols.forEach((col, idx) => {
      const val = normalizeAvailabilityLiteral_(row[col]);
      if (val) {
        availability[name][idx + 1] = val;
      }
    });
  }
  return availability;
}

/**
 * ShiftPattern 読み込み
 */
function readShiftPatterns_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.sheets.patterns);
  if (!sheet) {
    throw new Error(`ShiftPattern シート ${CONFIG.sheets.patterns} が見つかりません`);
  }
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    throw new Error('ShiftPattern シートにデータがありません');
  }
  const header = values[0];
  let codeCol = findHeaderIndexFlexible_(header, CONFIG.shiftPatternColHints.code);
  if (codeCol < 0) {
    throw new Error('ShiftPattern シートに「コード」列が見つかりません');
  }
  
  // 「労働時間」列のインデックスを取得
  const hoursCol = findHeaderIndexFlexible_(header, CONFIG.shiftPatternColHints.hours);
  if (hoursCol < 0) {
    throw new Error('ShiftPattern シートに「労働時間」列が見つかりません');
  }

  const metaLabels = new Set([
    'コード',
    '区別',
    '拘束時間',
    '休憩',
    '労働時間',
    'StaffDB_Column',
    ...CONFIG.shiftPatternColHints.code,
    ...CONFIG.shiftPatternColHints.department,
    ...CONFIG.shiftPatternColHints.hours
  ]);

  const departmentColumns = header
    .map((title, idx) => ({ title: title ? title.toString().trim() : '', idx }))
    .filter(item => item.idx !== codeCol && item.title && !metaLabels.has(item.title));

  if (!departmentColumns.length) {
    throw new Error('ShiftPattern シートに部門列が見つかりません（朝食・ロビー等の列が必要です）');
  }

  const result = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const code = row[codeCol];
    if (!code) continue;
    if (!result[code]) {
      result[code] = {
        code,
        cover: {},
        hours: [],
        raw: row,
        hoursColIndex: hoursCol // 労働時間列のインデックスを保存
      };
    }

    departmentColumns.forEach(col => {
      const cell = row[col.idx];
      if (!cell) {
        return;
      }
      const hourSet = parsePatternHours_(cell, cell);
      if (!hourSet.size) {
        return;
      }
      if (!result[code].cover[col.title]) {
        result[code].cover[col.title] = [];
      }
      result[code].cover[col.title] = union_(result[code].cover[col.title], Array.from(hourSet));
      result[code].hours = union_(result[code].hours, Array.from(hourSet));
    });
  }
  return result;
}

/**
 * Demand 読み込み
 */
function readDemand_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.sheets.demand);
  if (!sheet) {
    throw new Error(`Demand シート ${CONFIG.sheets.demand} が見つかりません`);
  }
  const values = sheet.getDataRange().getDisplayValues();
  const headerRowIndex = findHeaderRowIndex_(values, ['部門', '項目', 'Department'], ['時間', '時間帯', 'HourRange', 'Time']);
  if (headerRowIndex < 0) {
    throw new Error('Demand シートに有効なヘッダ行が見つかりません (部門/項目 と 時間 列が必要)');
  }
  const header = values[headerRowIndex];
  const dayCols = extractDayColumns_(header);
  const departmentCol = findHeaderIndexFlexible_(header, ['部門', '部門名', 'Department', '項目']);
  const timeCol = findHeaderIndexFlexible_(header, ['時間帯', 'HourRange', '時間', 'Time']);
  if (departmentCol < 0) {
    throw new Error(`Demand シートに「部門/項目」列が見つかりません: header=${JSON.stringify(header)}`);
  }
  if (timeCol < 0) {
    throw new Error(`Demand シートに「時間」列が見つかりません: header=${JSON.stringify(header)}`);
  }

  const demand = {};
  for (let day = 1; day <= DAY_COUNT; day++) {
    demand[day] = {};
  }

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const department = row[departmentCol];
    const rangeStr = row[timeCol];
    const hours = expandTimeRange_(rangeStr);
    dayCols.forEach((col, idx) => {
      const need = Number(row[col]) || 0;
      const day = idx + 1;
      if (!demand[day][department]) {
        demand[day][department] = {};
      }
      hours.forEach(hour => {
        demand[day][department][hour] = need;
      });
    });
  }
  return demand;
}

/**
 * assignments 初期化
 */
function initAssignments_(staff) {
  const assignments = {};
  staff.forEach(st => {
    assignments[st.name] = {
      days: Array(DAY_COUNT).fill('')
    };
  });
  return assignments;
}

/**
 * metrics 初期化
 */
function initMetrics_(staff) {
  const metrics = {};
  staff.forEach(st => {
    metrics[st.name] = {
      totalHours: 0,
      workDays: 0,
      holidayCount: 0,
      consecutive: 0
    };
  });
  return metrics;
}

/**
 * 指定スタッフが指定日にコードを配置できるか判定
 */
function canAssignPattern_(ctx, staff, day, code, options) {
  const pattern = ctx.patterns[code];
  if (!pattern) {
    return false;
  }
  if (!staff.shiftFlags[code]) {
    return false;
  }
  if (!isDayAvailable_(ctx, staff.name, day)) {
    return false;
  }
  if (ctx.assignments[staff.name].days[day - 1]) {
    return false;
  }
  if (!hasRemainingDemandForPattern_(ctx, day, pattern)) {
    return false;
  }
  if (wouldBreakConsecutiveLimit_(ctx, staff, day)) {
    return false;
  }
  if (wouldExceedMaxHours_(ctx, staff, pattern)) {
    return false;
  }
  return true;
}

/**
 * Demand は無視して「連勤・MaxHours・可否・NG/OFF/公休」のみを見る版
 * （既に別の人が入っていたシフトコードを他の人に付け替えるとき用）
 */
function canAssignPatternIgnoringDemand_(ctx, staff, day, code) {
  const pattern = ctx.patterns[code];
  if (!pattern) {
    return false;
  }
  if (!staff.shiftFlags[code]) {
    return false;
  }
  if (!isDayAvailable_(ctx, staff.name, day)) {
    return false;
  }
  if (ctx.assignments[staff.name].days[day - 1]) {
    return false;
  }
  if (wouldBreakConsecutiveLimit_(ctx, staff, day)) {
    return false;
  }
  if (wouldExceedMaxHours_(ctx, staff, pattern)) {
    return false;
  }
  return true;
}

function commitAssignment_(ctx, staff, day, code) {
  if (!ctx.assignments[staff.name]) {
    Logger.log(`⚠️ エラー: ${staff.name} の assignment が存在しません`);
    return;
  }
  const oldCode = ctx.assignments[staff.name].days[day - 1];
  ctx.assignments[staff.name].days[day - 1] = code;
  ctx.logs.push(`assign day${day} ${code} -> ${staff.name}`);
  if (oldCode !== code) {
    Logger.log(`  commitAssignment: ${staff.name} の ${day}日目 = ${oldCode} -> ${code}`);
  }
}

function isDayAvailable_(ctx, staffName, day) {
  const availability = ctx.availability[staffName]?.[day];
  if (!availability) return true;
  return ![CONFIG.literals.off, CONFIG.literals.ng, CONFIG.literals.holiday].includes(availability);
}

function hasRemainingDemandForPattern_(ctx, day, pattern) {
  if (!pattern || !pattern.cover) {
    return true;
  }
  for (const department of Object.keys(pattern.cover)) {
    const hours = pattern.cover[department] || [];
    for (const hour of hours) {
      const demandVal = ctx.demand[day]?.[department]?.[hour] || 0;
      const haveVal = calcHave_(ctx, day, department, hour);
      if (haveVal < demandVal) {
        return true;
      }
    }
  }
  return false;
}

function wouldExceedMaxHours_(ctx, staff, pattern) {
  if (!staff.maxHours) {
    return false;
  }
  const current = getAssignedHoursForStaff_(ctx, staff.name);
  const next = getPatternHourCount_(pattern);
  return current + next > staff.maxHours;
}

function getAssignedHoursForStaff_(ctx, staffName) {
  const assignment = ctx.assignments[staffName];
  if (!assignment) return 0;
  return assignment.days.reduce((sum, code) => {
    if (!code) return sum;
    const pattern = ctx.patterns[code];
    if (!pattern) return sum;
    return sum + getPatternHourCount_(pattern);
  }, 0);
}


function wouldBreakConsecutiveLimit_(ctx, staff, day) {
  const limit = CONFIG.limits.maxConsecutive; // 正社員・パート共通で5日
  const assignment = ctx.assignments[staff.name];
  if (!assignment) return false;
  let streak = 0;
  for (let d = day - 2; d >= 0; d--) {
    const code = assignment.days[d];
    if (code && code !== CONFIG.literals.holiday && code !== CONFIG.literals.off && code !== CONFIG.literals.ng) {
      streak += 1;
    } else {
      break;
    }
  }
  return streak + 1 > limit;
}

function isCookSpecialist_(staff, cookCodes) {
  const enabledCodes = Object.keys(staff.shiftFlags).filter(code => staff.shiftFlags[code]);
  if (!enabledCodes.length) return false;
  return enabledCodes.every(code => cookCodes.includes(code));
}

/**
 * metrics 再計算
 */
function recalcMetrics_(ctx) {
  Object.keys(ctx.metrics).forEach(name => {
    const metric = ctx.metrics[name];
    const assignment = ctx.assignments[name];
    const staff = ctx.staff.find(st => st.name === name);
    metric.totalHours = 0;
    metric.workDays = 0;
    metric.holidayCount = 0;
    metric.consecutive = 0;

    assignment.days.forEach(code => {
      if (!code) {
        // 空欄は公休としてカウント
        metric.consecutive = 0;
        metric.holidayCount += 1;
        return;
      }
      if (code === CONFIG.literals.off) {
        // OFF は有給（公休とは別カウント、公休数には含めない）
        metric.consecutive = 0;
        return;
      }
      if (code === CONFIG.literals.ng || code === CONFIG.literals.holiday) {
        // NG と 公休 は公休としてカウント（NG含めて最大4）
        metric.holidayCount += 1;
        metric.consecutive = 0;
        return;
      }
      const pattern = ctx.patterns[code];
      if (!pattern) {
        return;
      }
      // 労働時間は ShiftPattern の「労働時間」列から取得
      const hours = getPatternHourCount_(pattern);
      metric.totalHours += hours;
      metric.workDays += 1;
      metric.consecutive += 1;
    });
  });
}

/**
 * need 完全一致チェック
 */
function validateNeedCoverage_(ctx) {
  for (let day = 1; day <= DAY_COUNT; day++) {
    const departments = Object.keys(ctx.demand[day]);
    departments.forEach(dep => {
      HOURS_SEQUENCE.forEach(hour => {
        const need = ctx.demand[day][dep]?.[hour] || 0;
        const have = calcHave_(ctx, day, dep, hour);
        if (need !== have) {
          ctx.warnings.push(`need mismatch day=${day} dep=${dep} hour=${hour} need=${need} have=${have}`);
        }
      });
    });
  }
}

/**
 * 労働時間/休日チェック
 */
function validateHourAndHoliday_(ctx) {
  ctx.staff.forEach(staff => {
    const metric = ctx.metrics[staff.name];
    if (metric.totalHours > staff.maxHours) {
      ctx.warnings.push(`${staff.name} が MaxHours (${staff.maxHours}) を超過: ${metric.totalHours}`);
    }
    if (staff.isFullTime && metric.holidayCount !== CONFIG.limits.defaultHolidayCount) {
      ctx.warnings.push(`${staff.name} の公休日数(${metric.holidayCount})が期待値と違います`);
    }
  });
}

/**
 * すべての制約をチェック（完全一致が必要）
 * @returns {string[]} 制約違反のリスト
 */
function checkAllConstraints_(ctx) {
  const violations = [];
  
  // 1. Demand の完全一致チェック
  for (let day = 1; day <= DAY_COUNT; day++) {
    const departments = Object.keys(ctx.demand[day]);
    departments.forEach(dep => {
      HOURS_SEQUENCE.forEach(hour => {
        const need = ctx.demand[day][dep]?.[hour] || 0;
        const have = calcHave_(ctx, day, dep, hour);
        if (need !== have) {
          violations.push(`Demand不一致 day=${day} dep=${dep} hour=${hour} need=${need} have=${have}`);
        }
      });
    });
  }
  
  // 2. 総労働時間の完全一致チェック（Min/Maxの間でなければならない）
  ctx.staff.forEach(staff => {
    const metric = ctx.metrics[staff.name];
    if (metric.totalHours < staff.minHours) {
      violations.push(`${staff.name} 総労働時間がMin未達: ${metric.totalHours} < ${staff.minHours}`);
    }
    if (metric.totalHours > staff.maxHours) {
      violations.push(`${staff.name} 総労働時間がMax超過: ${metric.totalHours} > ${staff.maxHours}`);
    }
  });
  
  // 3. 公休数の完全一致チェック（正社員は4日）
  ctx.staff.forEach(staff => {
    if (staff.isFullTime) {
      const metric = ctx.metrics[staff.name];
      if (metric.holidayCount !== CONFIG.limits.defaultHolidayCount) {
        violations.push(`${staff.name} 公休数不一致: ${metric.holidayCount} !== ${CONFIG.limits.defaultHolidayCount}`);
      }
    }
  });
  
  return violations;
}

/**
 * ヘッダ文字列からヘッダマップを生成
 */
function createHeaderIndex_(header) {
  const map = {};
  Object.keys(CONFIG.colHints).forEach(key => {
    map[key] = findFirstIndex_(header, CONFIG.colHints[key]);
  });
  return map;
}

function findFirstIndex_(header, candidates) {
  for (let i = 0; i < header.length; i++) {
    const text = header[i].trim();
    if (candidates.some(candidate => text === candidate)) {
      return i;
    }
  }
  return -1;
}

function normalizeBoolean_(value) {
  if (typeof value === 'boolean') return value;
  if (!value) return false;
  const normalized = value.toString().trim().toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === '○';
}

function parseNumber_(value) {
  if (typeof value === 'number') return value;
  if (!value) return NaN;
  return Number(value.toString().replace(/[^\d.-]/g, ''));
}

function extractDayColumns_(header) {
  const columns = [];
  header.forEach((title, idx) => {
    const normalized = title.replace('日', '').trim();
    const day = Number(normalized);
    if (!isNaN(day) && day >= 1 && day <= DAY_COUNT) {
      columns.push(idx);
    }
  });
  return columns;
}

function normalizeAvailabilityLiteral_(value) {
  if (!value) return '';
  const normalized = value.toString().trim().toUpperCase();
  if ([CONFIG.literals.ng, CONFIG.literals.off, CONFIG.literals.holiday].includes(normalized)) {
    return normalized;
  }
  return '';
}

function findHeaderIndexFlexible_(header, candidates) {
  if (!candidates || !candidates.length) {
    return -1;
  }
  const lowerCandidates = candidates.map(target => target.toString().trim().toLowerCase());
  for (let i = 0; i < header.length; i++) {
    const cell = header[i]?.toString().trim();
    if (!cell) continue;
    const lower = cell.toLowerCase();
    if (lowerCandidates.includes(lower)) {
      return i;
    }
  }
  for (let i = 0; i < header.length; i++) {
    const cell = header[i]?.toString().trim();
    if (!cell) continue;
    const lower = cell.toLowerCase();
    if (lowerCandidates.some(target => lower.indexOf(target) >= 0)) {
      return i;
    }
  }
  return -1;
}

function findHeaderRowIndex_(rows, ...labelGroups) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;
    const ok = labelGroups.every(group => !group || !group.length || findHeaderIndexFlexible_(row, group) >= 0);
    if (ok) {
      return i;
    }
  }
  return -1;
}

function expandTimeRange_(rangeStr) {
  if (!rangeStr) return [];
  const normalized = rangeStr.toString().trim();
  if (!normalized) return [];
  const tokens = normalized.split(/-|〜|～|–|to/gi).map(token => token.trim()).filter(Boolean);
  if (!tokens.length) return [];

  const parseHour = value => {
    if (!value) return NaN;
    const match = value.match(/(\d{1,2})/);
    if (!match) return NaN;
    const hour = Number(match[1]);
    return isNaN(hour) ? NaN : ((hour % 24) + 24) % 24;
  };

  const start = parseHour(tokens[0]);
  let end = tokens.length > 1 ? parseHour(tokens[1]) : NaN;
  if (isNaN(start)) return [];
  if (isNaN(end)) {
    end = (start + 1) % 24;
  }

  const result = [];
  if (start === end) {
    result.push(start);
    return result;
  }
  if (start < end) {
    for (let h = start; h < end; h++) {
      result.push(h);
    }
  } else {
    for (let h = start; h < 24; h++) result.push(h);
    for (let h = 0; h < end; h++) result.push(h);
  }
  return result;
}

function buildHourSequence_() {
  const seq = [];
  for (let h = 5; h < 24; h++) seq.push(h);
  for (let h = 0; h < 5; h++) seq.push(h);
  return seq;
}

function union_(arr1, arr2) {
  const set = new Set([].concat(arr1 || [], arr2 || []));
  return Array.from(set).sort((a, b) => a - b);
}

function parsePatternHours_(listText, rangeText) {
  const hours = new Set();
  const pushList = text => {
    if (!text) return;
    text
      .split(/[,\/\s]+/)
      .map(token => Number(token))
      .filter(num => !isNaN(num))
      .forEach(num => hours.add(num));
  };
  pushList(listText);
  if (!hours.size && rangeText) {
    expandTimeRange_(rangeText).forEach(h => hours.add(h));
  }
  if (!hours.size && listText) {
    expandTimeRange_(listText).forEach(h => hours.add(h));
  }
  return hours;
}

/**
 * Debug 用
 */
function dumpContext() {
  const ctx = loadContext_();
  Logger.log(JSON.stringify(ctx, null, 2));
}

/**
 * TODO: 各フェーズのテスト関数
 */
function testNightPhase() {
  const ctx = loadContext_();
  phaseNight_(ctx);
  Logger.log(JSON.stringify(ctx.assignments, null, 2));
}


