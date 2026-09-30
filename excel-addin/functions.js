// 自動生成(excel-addin/scripts/build.mjs)。直接編集しない
(() => {
  var __defProp = Object.defineProperty;
  var __defProps = Object.defineProperties;
  var __getOwnPropDescs = Object.getOwnPropertyDescriptors;
  var __getOwnPropSymbols = Object.getOwnPropertySymbols;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __propIsEnum = Object.prototype.propertyIsEnumerable;
  var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
  var __spreadValues = (a, b) => {
    for (var prop in b || (b = {}))
      if (__hasOwnProp.call(b, prop))
        __defNormalProp(a, prop, b[prop]);
    if (__getOwnPropSymbols)
      for (var prop of __getOwnPropSymbols(b)) {
        if (__propIsEnum.call(b, prop))
          __defNormalProp(a, prop, b[prop]);
      }
    return a;
  };
  var __spreadProps = (a, b) => __defProps(a, __getOwnPropDescs(b));

  // ../job-extension/src/parser.js
  (() => {
    const DEFAULT_WORKDAYS_PER_MONTH = 20.4;
    const DEFAULT_HOURS_PER_DAY = 8;
    const toHalf = (s) => String(s != null ? s : "").normalize("NFKC");
    function parseYen(s) {
      var _a;
      const t = toHalf(s).replace(/,/g, "");
      const m = t.match(/(\d+(?:\.\d+)?)\s*万\s*(\d+)?\s*円?|(\d+(?:\.\d+)?)\s*円/);
      if (!m) return null;
      if (m[1] != null) return Math.round(Number(m[1]) * 1e4 + Number((_a = m[2]) != null ? _a : 0));
      return Math.round(Number(m[3]));
    }
    const YEN = String.raw`\d+(?:[.,]\d+)*\s*万\s*(?:\d+(?:,\d+)*\s*)?円?|\d+(?:,\d+)*\s*円`;
    const TILDE = String.raw`\s*[~〜～\-－ー]\s*`;
    function findRange(text2, kinds, gap = 6) {
      const re = new RegExp(`(${kinds})[^\\d]{0,${gap}}(${YEN})(?:${TILDE}(${YEN}))?`);
      const m = text2.match(re);
      if (!m) return null;
      const min = parseYen(m[2]);
      const max = m[3] ? parseYen(m[3]) : null;
      if (min == null) return null;
      const validMax = max != null && max > min ? max : null;
      const rest = text2.slice(m.index + m[0].length);
      const open = validMax == null && max == null && /^\s*(?:[~〜～]|以上)/.test(rest);
      return { kind: m[1], min, max: validMax, open };
    }
    function findBonusMonths(text2) {
      const m = text2.match(/賞与[^。]{0,40}?(\d+(?:\.\d+)?)\s*[ヶヵかカケ箇]?\s*月分/);
      return m ? Number(m[1]) : null;
    }
    function findFixedOvertime(text2) {
      if (/(固定残業|みなし残業)[^。\n]{0,10}(ありません|なし|無し|ございません)/.test(text2)) {
        return null;
      }
      const idx = text2.search(/固定残業|みなし残業/);
      if (idx < 0) return null;
      const seg = text2.slice(idx, idx + 80);
      const hours = seg.match(/(\d+(?:\.\d+)?)\s*時間/);
      const amount = seg.match(new RegExp(YEN));
      return {
        hours: hours ? Number(hours[1]) : null,
        amount: amount ? parseYen(amount[0]) : null
      };
    }
    function findHoursPerDay(text2) {
      const m = text2.match(
        /(?:実働|所定労働時間|労働時間)[^\d]{0,4}(?:1日(?:あたり|当たり|当り)?[^\d]{0,2})?(\d+(?:\.\d+)?)\s*時間\s*(?:(\d+)\s*分)?/
      );
      if (m) {
        const h = Number(m[1]) + (m[2] ? Number(m[2]) / 60 : 0);
        return h > 0 && h <= 12 ? h : null;
      }
      return hoursFromSpan(text2);
    }
    function hoursFromSpan(text2) {
      var _a, _b;
      const span = text2.match(
        /(\d{1,2})\s*[時:]\s*(\d{2})?\s*分?\s*[~〜～\-－]\s*(\d{1,2})\s*[時:]\s*(\d{2})?\s*分?/
      );
      const rest = text2.match(/休憩(?:時間)?[^\d]{0,6}(\d+(?:\.\d+)?)\s*(分|時間)/);
      if (!span || !rest) return null;
      const start = Number(span[1]) + Number((_a = span[2]) != null ? _a : 0) / 60;
      const end = Number(span[3]) + Number((_b = span[4]) != null ? _b : 0) / 60;
      const breakHours = Number(rest[1]) / (rest[2] === "分" ? 60 : 1);
      const h = end - start - breakHours;
      return h > 0 && h <= 12 ? Math.round(h * 100) / 100 : null;
    }
    function findHolidays(text2) {
      const m = text2.match(/年間休日[^\d]{0,4}(\d{2,3})\s*日/);
      const n = m ? Number(m[1]) : null;
      return n != null && n >= 50 && n <= 200 ? n : null;
    }
    const STATED_KINDS = [
      "想定年収",
      "初年度(?:の|想定)?年収",
      "年収例",
      "(?<!平均)年収",
      "平均年収"
    ];
    function findStatedAnnual(text2) {
      for (const kind of STATED_KINDS) {
        const r = findRange(text2, kind, 25);
        if (r && r.min >= 1e6) return __spreadProps(__spreadValues({}, r), { companyAverage: kind === "平均年収" });
      }
      return null;
    }
    function analyze(salaryText, workText = "") {
      var _a;
      const salary = toHalf(salaryText);
      const work = toHalf(workText);
      const all = `${salary}
${work}`;
      const monthly = findRange(salary, "月給|月収|基本給|月額");
      const hourly = findRange(salary, "時給(?!換算)");
      const daily = findRange(salary, "日給");
      const yearly = findRange(salary, "年俸");
      const stated = findStatedAnnual(salary);
      const bonusMonths = findBonusMonths(salary);
      const fixedOvertime = findFixedOvertime(all);
      const hoursPerDay = (_a = findHoursPerDay(work)) != null ? _a : findHoursPerDay(salary);
      const holidays = findHolidays(all);
      const workdays = holidays != null ? (365 - holidays) / 12 : DEFAULT_WORKDAYS_PER_MONTH;
      const hours = hoursPerDay != null ? hoursPerDay : DEFAULT_HOURS_PER_DAY;
      const monthlyHours = workdays * hours;
      const assumptions = [];
      if (holidays == null) assumptions.push("年間休日120日");
      if (hoursPerDay == null) assumptions.push("1日8時間");
      let base = null;
      let basis = null;
      if (monthly) {
        base = monthly;
        basis = monthly.kind;
      } else if (daily) {
        base = {
          min: daily.min * workdays,
          max: daily.max && daily.max * workdays,
          open: daily.open
        };
        basis = "日給";
      } else if (hourly) {
        base = {
          min: hourly.min * monthlyHours,
          max: hourly.max && hourly.max * monthlyHours,
          open: hourly.open
        };
        basis = "時給";
      }
      let annual = null;
      if (yearly && yearly.min >= 1e6) {
        annual = {
          min: yearly.min,
          max: yearly.max,
          open: yearly.open,
          bonusIncluded: true,
          from: "年俸"
        };
      } else if (base) {
        const months = 12 + (bonusMonths != null ? bonusMonths : 0);
        annual = {
          min: Math.round(base.min * months),
          max: base.max ? Math.round(base.max * months) : null,
          open: Boolean(base.open),
          bonusIncluded: bonusMonths != null,
          from: basis
        };
      }
      let hourlyEquivalent = null;
      if (base && basis !== "時給") {
        hourlyEquivalent = {
          min: Math.round(base.min / monthlyHours),
          max: base.max ? Math.round(base.max / monthlyHours) : null
        };
      }
      let baseWithoutOvertime = null;
      if (base && (fixedOvertime == null ? void 0 : fixedOvertime.amount) && fixedOvertime.amount < base.min) {
        baseWithoutOvertime = {
          monthly: base.min - fixedOvertime.amount,
          hourly: Math.round((base.min - fixedOvertime.amount) / monthlyHours)
        };
      }
      const found = Boolean(annual || stated);
      return {
        found,
        basis,
        monthly: base && { min: Math.round(base.min), max: base.max ? Math.round(base.max) : null },
        bonusMonths,
        annual,
        stated,
        hourlyEquivalent,
        fixedOvertime,
        baseWithoutOvertime,
        hoursPerDay,
        holidays,
        assumptions
      };
    }
    globalThis.JobParser = { parseYen, analyze };
  })();

  // ../realty-extension/src/parser.js
  (() => {
    const TSUBO_PER_M2 = 0.3025;
    const toHalf = (s) => String(s != null ? s : "").normalize("NFKC");
    function parseYen(s) {
      var _a, _b, _c;
      const t = toHalf(s).replace(/,/g, "");
      const m = t.match(/(?:(\d+(?:\.\d+)?)\s*億)?\s*(?:(\d+(?:\.\d+)?)\s*万)?\s*(\d+)?\s*円/);
      if (!m || m[1] == null && m[2] == null && m[3] == null) return null;
      return Math.round(Number((_a = m[1]) != null ? _a : 0) * 1e8 + Number((_b = m[2]) != null ? _b : 0) * 1e4 + Number((_c = m[3]) != null ? _c : 0));
    }
    function parseArea(s) {
      const t = toHalf(s).replace(/,/g, "");
      const m2 = t.match(/(\d+(?:\.\d+)?)\s*(?:m2|m²|㎡|平米|平方メートル)/i);
      if (m2) return Number(m2[1]);
      const tsubo = t.match(/(\d+(?:\.\d+)?)\s*坪/);
      return tsubo ? Number(tsubo[1]) / TSUBO_PER_M2 : null;
    }
    function parsePercent(s) {
      const m = toHalf(s).match(/(\d+(?:\.\d+)?)\s*%/);
      return m ? Number(m[1]) / 100 : null;
    }
    function parseMonthsOrYen(s, rent) {
      const t = toHalf(s).trim();
      if (!t || /^[-ー―]$|なし|無し/.test(t)) return 0;
      const months = t.match(/(\d+(?:\.\d+)?)\s*[ヶヵかカケ箇]?\s*月/);
      if (months && !/円/.test(t)) return rent != null ? Math.round(Number(months[1]) * rent) : null;
      return parseYen(t);
    }
    function monthlyPayment(principal, annualRate, years) {
      const n = years * 12;
      const r = annualRate / 12;
      if (r === 0) return principal / n;
      return principal * r / (1 - (1 + r) ** -n);
    }
    function analyze(fields, options2 = {}) {
      var _a;
      const { loanRate = 0.01, loanYears = 35 } = options2;
      const f = (k) => {
        var _a2;
        return (_a2 = fields[k]) != null ? _a2 : null;
      };
      const area = f("area") != null ? parseArea(f("area")) : null;
      const tsubo = area != null ? area * TSUBO_PER_M2 : null;
      const assumptions = [];
      const rent = f("rent") != null ? parseYen(f("rent")) : null;
      if (rent != null && rent > 0 && rent < 1e7) {
        const fee2 = f("fee") != null ? (_a = parseYen(f("fee"))) != null ? _a : 0 : 0;
        const monthly = rent + fee2;
        let deposit = null;
        let keyMoney = null;
        if (f("depositKey") != null) {
          const [d, k] = toHalf(f("depositKey")).split(/[/／]/);
          deposit = parseMonthsOrYen(d, rent);
          keyMoney = k != null ? parseMonthsOrYen(k, rent) : null;
        }
        if (f("deposit") != null) deposit = parseMonthsOrYen(f("deposit"), rent);
        if (f("keyMoney") != null) keyMoney = parseMonthsOrYen(f("keyMoney"), rent);
        let initialCost = null;
        if (deposit != null && keyMoney != null) {
          initialCost = Math.round(deposit + keyMoney + rent * 1.1 + monthly);
          assumptions.push("仲介手数料は賃料1ヶ月分+税、前家賃1ヶ月分で計算");
        }
        return {
          kind: "rent",
          rent,
          fee: fee2,
          monthly,
          area,
          perM2: area ? Math.round(monthly / area) : null,
          perTsubo: tsubo ? Math.round(monthly / tsubo) : null,
          deposit,
          keyMoney,
          initialCost,
          assumptions
        };
      }
      const price = f("price") != null ? parseYen(f("price")) : null;
      if (price == null || price < 1e6) return null;
      let fee = null;
      let repair = null;
      if (f("feeRepair") != null) {
        const [a, b] = toHalf(f("feeRepair")).split(/[/／]/);
        fee = parseYen(a);
        repair = b != null ? parseYen(b) : null;
      }
      if (f("fee") != null) fee = parseYen(f("fee"));
      if (f("repair") != null) repair = parseYen(f("repair"));
      const runningMonthly = (fee != null ? fee : 0) + (repair != null ? repair : 0);
      const payment = Math.round(monthlyPayment(price, loanRate, loanYears));
      assumptions.push(
        `全額を金利${Math.round(loanRate * 1e4) / 100}%・${loanYears}年の元利均等で借りた場合`
      );
      let grossYield = f("yield") != null ? parsePercent(f("yield")) : null;
      let annualIncome = f("income") != null ? parseYen(f("income")) : null;
      if (annualIncome != null && grossYield == null) grossYield = annualIncome / price;
      if (grossYield != null && annualIncome == null) annualIncome = Math.round(price * grossYield);
      const netYield = annualIncome != null && runningMonthly > 0 ? (annualIncome - runningMonthly * 12) / price : null;
      return {
        kind: "sale",
        price,
        area,
        perM2: area ? Math.round(price / area) : null,
        perTsubo: tsubo ? Math.round(price / tsubo) : null,
        fee,
        repair,
        loanPayment: payment,
        monthlyTotal: payment + runningMonthly,
        grossYield,
        annualIncome,
        netYield,
        assumptions
      };
    }
    globalThis.RealtyParser = { parseYen, parseArea, parsePercent, monthlyPayment, analyze };
  })();

  // ../calc-api/src/calc.js
  var Job = globalThis.JobParser;
  var Realty = globalThis.RealtyParser;
  var MAX_TEXT = 2e3;
  var InputError = class extends Error {
    constructor(message, field) {
      super(message);
      this.field = field;
    }
  };
  var BASIS = {
    月給: "monthly_salary",
    月収: "monthly_income",
    基本給: "base_salary",
    月額: "monthly_amount",
    日給: "daily_wage",
    時給: "hourly_wage",
    年俸: "annual_salary"
  };
  var STATED = {
    想定年収: "expected",
    年収例: "example",
    年収: "stated",
    平均年収: "company_average"
  };
  var assumption = (code, ja, en) => ({ code, ja, en });
  var JOB_ASSUMPTIONS = {
    年間休日120日: assumption(
      "holidays_default_120",
      "年間休日120日",
      "Assumed 120 annual holidays (not stated)"
    ),
    "1日8時間": assumption(
      "hours_default_8",
      "1日8時間",
      "Assumed 8 working hours per day (not stated)"
    )
  };
  function text(body, field, { required = false } = {}) {
    const v = body == null ? void 0 : body[field];
    if (v == null || v === "") {
      if (required) throw new InputError(`"${field}" is required`, field);
      return "";
    }
    if (typeof v !== "string") throw new InputError(`"${field}" must be a string`, field);
    if (v.length > MAX_TEXT) throw new InputError(`"${field}" is too long (max ${MAX_TEXT})`, field);
    return v;
  }
  var range = (r) => {
    var _a;
    return r ? { min: r.min, max: (_a = r.max) != null ? _a : null } : null;
  };
  function analyzeSalary(body) {
    var _a, _b, _c, _d;
    const salary = text(body, "salary", { required: true });
    const work = text(body, "workConditions");
    const r = Job.analyze(salary, work);
    return {
      found: r.found,
      basis: r.basis ? { code: (_a = BASIS[r.basis]) != null ? _a : "other", ja: r.basis } : null,
      monthly: range(r.monthly),
      bonusMonths: r.bonusMonths,
      annual: r.annual && {
        min: r.annual.min,
        max: (_b = r.annual.max) != null ? _b : null,
        openEnded: r.annual.open,
        bonusIncluded: r.annual.bonusIncluded
      },
      statedAnnual: r.stated && {
        type: (_c = STATED[r.stated.kind]) != null ? _c : "stated",
        ja: r.stated.kind,
        min: r.stated.min,
        max: (_d = r.stated.max) != null ? _d : null,
        openEnded: r.stated.open,
        companyAverage: r.stated.companyAverage
      },
      hourlyEquivalent: range(r.hourlyEquivalent),
      fixedOvertime: r.fixedOvertime && {
        hours: r.fixedOvertime.hours,
        amount: r.fixedOvertime.amount
      },
      withoutFixedOvertime: r.baseWithoutOvertime && {
        monthly: r.baseWithoutOvertime.monthly,
        hourly: r.baseWithoutOvertime.hourly
      },
      hoursPerDay: r.hoursPerDay,
      annualHolidays: r.holidays,
      assumptions: r.assumptions.map((a) => {
        var _a2;
        return (_a2 = JOB_ASSUMPTIONS[a]) != null ? _a2 : assumption("other", a, a);
      }),
      currency: "JPY"
    };
  }

  // ../calc-api/src/takehome.js
  var RATES = {
    2026: {
      // 健康保険料率(%・労使合計)。令和8年3月分から
      health: {
        北海道: 10.28,
        青森: 9.85,
        岩手: 9.51,
        宮城: 10.1,
        秋田: 10.01,
        山形: 9.75,
        福島: 9.5,
        茨城: 9.52,
        栃木: 9.82,
        群馬: 9.68,
        埼玉: 9.67,
        千葉: 9.73,
        東京: 9.85,
        神奈川: 9.92,
        新潟: 9.21,
        富山: 9.59,
        石川: 9.7,
        福井: 9.71,
        山梨: 9.55,
        長野: 9.63,
        岐阜: 9.8,
        静岡: 9.61,
        愛知: 9.93,
        三重: 9.77,
        滋賀: 9.88,
        京都: 9.89,
        大阪: 10.13,
        兵庫: 10.12,
        奈良: 9.91,
        和歌山: 10.06,
        鳥取: 9.86,
        島根: 9.94,
        岡山: 10.05,
        広島: 9.78,
        山口: 10.15,
        徳島: 10.24,
        香川: 10.02,
        愛媛: 9.98,
        高知: 10.05,
        福岡: 10.11,
        佐賀: 10.55,
        長崎: 10.06,
        熊本: 10.08,
        大分: 10.08,
        宮崎: 9.77,
        鹿児島: 10.13,
        沖縄: 9.44
      },
      nursingCare: 1.62,
      // 介護保険(40〜64歳)
      childSupport: 0.23,
      // 子ども・子育て支援金(令和8年4月分から)
      pension: 18.3,
      // 厚生年金
      employment: 0.5,
      // 雇用保険・労働者負担(一般の事業)
      pensionCap: 65e4,
      // 厚生年金の標準報酬月額の上限
      bonusCapPension: 15e5,
      // 厚生年金の標準賞与額の上限(1回あたり)
      bonusCapHealth: 573e4,
      // 健康保険の標準賞与額の上限(年度あたり)
      // 所得税の基礎控除(合計所得金額の上限, 控除額)。令和8・9年分の特例加算を含む
      basicDeduction: [
        [489e4, 104e4],
        [655e4, 67e4],
        [235e5, 62e4],
        [24e6, 48e4],
        [245e5, 32e4],
        [25e6, 16e4]
      ],
      residentBasicDeduction: 43e4,
      residentPerCapita: 5e3
      // 均等割(市町村3,000+道府県1,000+森林環境税1,000)
    }
  };
  var DEFAULT_YEAR = 2026;
  var ROMAJI = {
    hokkaido: "北海道",
    aomori: "青森",
    iwate: "岩手",
    miyagi: "宮城",
    akita: "秋田",
    yamagata: "山形",
    fukushima: "福島",
    ibaraki: "茨城",
    tochigi: "栃木",
    gunma: "群馬",
    saitama: "埼玉",
    chiba: "千葉",
    tokyo: "東京",
    kanagawa: "神奈川",
    niigata: "新潟",
    toyama: "富山",
    ishikawa: "石川",
    fukui: "福井",
    yamanashi: "山梨",
    nagano: "長野",
    gifu: "岐阜",
    shizuoka: "静岡",
    aichi: "愛知",
    mie: "三重",
    shiga: "滋賀",
    kyoto: "京都",
    osaka: "大阪",
    hyogo: "兵庫",
    nara: "奈良",
    wakayama: "和歌山",
    tottori: "鳥取",
    shimane: "島根",
    okayama: "岡山",
    hiroshima: "広島",
    yamaguchi: "山口",
    tokushima: "徳島",
    kagawa: "香川",
    ehime: "愛媛",
    kochi: "高知",
    fukuoka: "福岡",
    saga: "佐賀",
    nagasaki: "長崎",
    kumamoto: "熊本",
    oita: "大分",
    miyazaki: "宮崎",
    kagoshima: "鹿児島",
    okinawa: "沖縄"
  };
  function prefectureKey(input) {
    var _a, _b;
    const s = String(input != null ? input : "").normalize("NFKC").trim();
    if (!s) return null;
    const ja = s === "北海道" ? s : s.replace(/[都府県]$/, "");
    if (RATES[DEFAULT_YEAR].health[ja] != null) return ja;
    const lower = s.toLowerCase().replace(/\s+/g, " ");
    const en = lower.replace(/[-\s](to|fu|ken)$/, "").replace(/\s*prefecture$/, "").replace(/\s+/g, "");
    return (_b = (_a = ROMAJI[lower.replace(/\s+/g, "")]) != null ? _a : ROMAJI[en]) != null ? _b : null;
  }
  var GRADES = [
    [63e3, 58e3],
    [73e3, 68e3],
    [83e3, 78e3],
    [93e3, 88e3],
    [101e3, 98e3],
    [107e3, 104e3],
    [114e3, 11e4],
    [122e3, 118e3],
    [13e4, 126e3],
    [138e3, 134e3],
    [146e3, 142e3],
    [155e3, 15e4],
    [165e3, 16e4],
    [175e3, 17e4],
    [185e3, 18e4],
    [195e3, 19e4],
    [21e4, 2e5],
    [23e4, 22e4],
    [25e4, 24e4],
    [27e4, 26e4],
    [29e4, 28e4],
    [31e4, 3e5],
    [33e4, 32e4],
    [35e4, 34e4],
    [37e4, 36e4],
    [395e3, 38e4],
    [425e3, 41e4],
    [455e3, 44e4],
    [485e3, 47e4],
    [515e3, 5e5],
    [545e3, 53e4],
    [575e3, 56e4],
    [605e3, 59e4],
    [635e3, 62e4],
    [665e3, 65e4],
    [695e3, 68e4],
    [73e4, 71e4],
    [77e4, 75e4],
    [81e4, 79e4],
    [855e3, 83e4],
    [905e3, 88e4],
    [955e3, 93e4],
    [1005e3, 98e4],
    [1055e3, 103e4],
    [1115e3, 109e4],
    [1175e3, 115e4],
    [1235e3, 121e4],
    [1295e3, 127e4],
    [1355e3, 133e4],
    [Infinity, 139e4]
  ];
  function standardMonthly(pay) {
    return GRADES.find(([limit]) => pay < limit)[1];
  }
  var standardMonthlyPension = (pay, cap) => Math.min(Math.max(standardMonthly(pay), 88e3), cap);
  var employeeShare = (amount) => {
    const floor = Math.floor(amount);
    return amount - floor > 0.5 ? floor + 1 : floor;
  };
  function salaryIncomeForIncomeTax2026(gross) {
    if (gross < 691e3) return 0;
    if (gross < 2191e3) return gross - 74e4;
    if (gross < 2193e3) return 1451e3;
    if (gross < 2196e3) return 1453e3;
    if (gross < 22e5) return 1456e3;
    return salaryIncomeStandard(gross);
  }
  function salaryIncomeStandard(gross, minimumDeduction = 69e4) {
    if (gross >= 85e5) return gross - 195e4;
    if (gross >= 66e5) return Math.floor(gross * 0.9 - 11e5);
    const a = Math.floor(gross / 4e3) * 4e3;
    const byFormula = gross >= 36e5 ? a * 0.8 - 44e4 : gross >= 1628e3 ? a * 0.7 - 8e4 : null;
    const byMinimum = Math.max(gross - minimumDeduction, 0);
    return byFormula == null ? byMinimum : Math.min(Math.floor(byFormula), byMinimum);
  }
  var INCOME_TAX = [
    [195e4, 0.05, 0],
    [33e5, 0.1, 97500],
    [695e4, 0.2, 427500],
    [9e6, 0.23, 636e3],
    [18e6, 0.33, 1536e3],
    [4e7, 0.4, 2796e3],
    [Infinity, 0.45, 4796e3]
  ];
  function incomeTax(taxable) {
    const t = Math.floor(Math.max(taxable, 0) / 1e3) * 1e3;
    const [, rate, minus] = INCOME_TAX.find(([limit]) => t <= limit);
    const base = Math.max(t * rate - minus, 0);
    return Math.floor(base * 1.021 / 100) * 100;
  }
  function residentTax(totalIncome, socialInsurance, r) {
    if (totalIncome <= 45e4) return 0;
    const taxable = Math.floor(Math.max(totalIncome - socialInsurance - r.residentBasicDeduction, 0) / 1e3) * 1e3;
    const adjust = taxable <= 2e6 ? Math.min(5e4, taxable) * 0.05 : Math.max((5e4 - (taxable - 2e6)) * 0.05, 2500);
    const incomeLevy = Math.floor(Math.max(taxable * 0.1 - adjust, 0) / 100) * 100;
    return incomeLevy + r.residentPerCapita;
  }
  var assumption2 = (code, ja, en) => ({ code, ja, en });
  function num(body, key, { required = false, min = 0, max, def, integer = false } = {}) {
    const v = body == null ? void 0 : body[key];
    if (v == null) {
      if (required) throw new InputError(`"${key}" is required`, key);
      return def;
    }
    if (typeof v !== "number" || !Number.isFinite(v) || v < min || max != null && v > max || integer && !Number.isInteger(v)) {
      throw new InputError(
        `"${key}" must be ${integer ? "an integer" : "a number"} between ${min} and ${max}`,
        key
      );
    }
    return v;
  }
  function calculateTakeHome(body) {
    var _a, _b;
    const monthly = Math.round(
      num(body, "monthlySalary", { required: true, min: 1, max: 2e7 })
    );
    const bonus = Math.round(num(body, "annualBonus", { max: 2e8, def: 0 }));
    const bonusTimes = num(body, "bonusTimes", {
      min: bonus > 0 ? 1 : 0,
      max: 12,
      def: bonus > 0 ? 2 : 0,
      integer: true
    });
    const age = num(body, "age", { min: 15, max: 99, def: 30, integer: true });
    const year = num(body, "year", {
      min: DEFAULT_YEAR,
      max: DEFAULT_YEAR,
      def: DEFAULT_YEAR,
      integer: true
    });
    const pref = prefectureKey((_a = body == null ? void 0 : body.prefecture) != null ? _a : "東京");
    if (!pref)
      throw new InputError(
        '"prefecture" must be a Japanese prefecture (e.g. "東京都" or "Tokyo")',
        "prefecture"
      );
    const r = RATES[year];
    const care = age >= 40 && age < 65;
    const healthRate = r.health[pref] + (care ? r.nursingCare : 0);
    const stdHealth = standardMonthly(monthly);
    const stdPension = standardMonthlyPension(monthly, r.pensionCap);
    const m = {
      health: employeeShare(stdHealth * r.health[pref] / 200),
      nursingCare: care ? employeeShare(stdHealth * r.nursingCare / 200) : 0,
      childSupport: employeeShare(stdHealth * r.childSupport / 200),
      pension: employeeShare(stdPension * r.pension / 200),
      employment: employeeShare(monthly * r.employment / 100)
    };
    const b = { health: 0, nursingCare: 0, childSupport: 0, pension: 0, employment: 0 };
    let healthBonusLeft = r.bonusCapHealth;
    for (let i = 0; i < bonusTimes; i++) {
      const pay = Math.floor(bonus / bonusTimes);
      const std = Math.floor(pay / 1e3) * 1e3;
      const stdH = Math.min(std, healthBonusLeft);
      healthBonusLeft -= stdH;
      b.health += employeeShare(stdH * r.health[pref] / 200);
      b.nursingCare += care ? employeeShare(stdH * r.nursingCare / 200) : 0;
      b.childSupport += employeeShare(stdH * r.childSupport / 200);
      b.pension += employeeShare(Math.min(std, r.bonusCapPension) * r.pension / 200);
      b.employment += employeeShare(pay * r.employment / 100);
    }
    const si = Object.fromEntries(Object.keys(m).map((k) => [k, m[k] * 12 + b[k]]));
    const siTotal = Object.values(si).reduce((a, x) => a + x, 0);
    const gross = monthly * 12 + bonus;
    const salaryIncome = salaryIncomeForIncomeTax2026(gross);
    const basic = ((_b = r.basicDeduction.find(([limit]) => salaryIncome <= limit)) != null ? _b : [0, 0])[1];
    const itax = incomeTax(salaryIncome - siTotal - basic);
    const rtax = residentTax(salaryIncomeStandard(gross), siTotal, r);
    const takeHome2 = gross - siTotal - itax - rtax;
    return {
      year,
      prefecture: pref,
      grossAnnual: gross,
      socialInsurance: {
        healthInsurance: si.health,
        nursingCareInsurance: si.nursingCare,
        childSupportLevy: si.childSupport,
        pension: si.pension,
        employmentInsurance: si.employment,
        total: siTotal,
        standardMonthlyRemuneration: { health: stdHealth, pension: stdPension }
      },
      incomeTax: itax,
      residentTax: rtax,
      totalDeductions: siTotal + itax + rtax,
      takeHomeAnnual: takeHome2,
      takeHomeMonthlyAverage: Math.round(takeHome2 / 12),
      takeHomeRatioPercent: Math.round(takeHome2 / gross * 1e3) / 10,
      rates: {
        healthInsurancePercent: r.health[pref],
        nursingCarePercent: care ? r.nursingCare : 0,
        childSupportPercent: r.childSupport,
        pensionPercent: r.pension,
        employmentInsuranceEmployeePercent: r.employment,
        totalEmployeeSharePercent: Math.round(((healthRate + r.childSupport + r.pension) / 2 + r.employment) * 1e3) / 1e3
      },
      assumptions: [
        assumption2(
          "employee_kyokai_kenpo",
          "会社員(協会けんぽ・厚生年金・雇用保険の一般の事業)で、独身・扶養なしとして計算",
          "Company employee in Kyokai Kenpo health insurance, single with no dependents"
        ),
        assumption2(
          "full_year_rates",
          `${year}年度の保険料率を1年分に当てはめて計算(子ども・子育て支援金は実際には4月分から)`,
          `FY${year} insurance rates applied to the whole year (the child support levy actually starts in April)`
        ),
        assumption2(
          "resident_tax_estimate",
          "住民税は同じ年の所得にかかる翌年度分を概算(基礎控除43万円・均等割5,000円、調整控除は基礎控除分のみ)",
          "Resident tax is estimated on the same year's income (it is actually levied the following fiscal year)"
        ),
        ...bonus > 0 ? [
          assumption2(
            "bonus_split",
            `賞与は年${bonusTimes}回に均等に分けて計算`,
            `Bonus is split evenly into ${bonusTimes} payments`
          )
        ] : []
      ],
      currency: "JPY"
    };
  }

  // ../calc-api/src/holidays.json
  var holidays_default = { source: "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv", from: "1955-01-01", to: "2027-12-31", holidays: { "1955-01-01": "元日", "1955-01-15": "成人の日", "1955-03-21": "春分の日", "1955-04-29": "天皇誕生日", "1955-05-03": "憲法記念日", "1955-05-05": "こどもの日", "1955-09-24": "秋分の日", "1955-11-03": "文化の日", "1955-11-23": "勤労感謝の日", "1956-01-01": "元日", "1956-01-15": "成人の日", "1956-03-21": "春分の日", "1956-04-29": "天皇誕生日", "1956-05-03": "憲法記念日", "1956-05-05": "こどもの日", "1956-09-23": "秋分の日", "1956-11-03": "文化の日", "1956-11-23": "勤労感謝の日", "1957-01-01": "元日", "1957-01-15": "成人の日", "1957-03-21": "春分の日", "1957-04-29": "天皇誕生日", "1957-05-03": "憲法記念日", "1957-05-05": "こどもの日", "1957-09-23": "秋分の日", "1957-11-03": "文化の日", "1957-11-23": "勤労感謝の日", "1958-01-01": "元日", "1958-01-15": "成人の日", "1958-03-21": "春分の日", "1958-04-29": "天皇誕生日", "1958-05-03": "憲法記念日", "1958-05-05": "こどもの日", "1958-09-23": "秋分の日", "1958-11-03": "文化の日", "1958-11-23": "勤労感謝の日", "1959-01-01": "元日", "1959-01-15": "成人の日", "1959-03-21": "春分の日", "1959-04-10": "結婚の儀", "1959-04-29": "天皇誕生日", "1959-05-03": "憲法記念日", "1959-05-05": "こどもの日", "1959-09-24": "秋分の日", "1959-11-03": "文化の日", "1959-11-23": "勤労感謝の日", "1960-01-01": "元日", "1960-01-15": "成人の日", "1960-03-20": "春分の日", "1960-04-29": "天皇誕生日", "1960-05-03": "憲法記念日", "1960-05-05": "こどもの日", "1960-09-23": "秋分の日", "1960-11-03": "文化の日", "1960-11-23": "勤労感謝の日", "1961-01-01": "元日", "1961-01-15": "成人の日", "1961-03-21": "春分の日", "1961-04-29": "天皇誕生日", "1961-05-03": "憲法記念日", "1961-05-05": "こどもの日", "1961-09-23": "秋分の日", "1961-11-03": "文化の日", "1961-11-23": "勤労感謝の日", "1962-01-01": "元日", "1962-01-15": "成人の日", "1962-03-21": "春分の日", "1962-04-29": "天皇誕生日", "1962-05-03": "憲法記念日", "1962-05-05": "こどもの日", "1962-09-23": "秋分の日", "1962-11-03": "文化の日", "1962-11-23": "勤労感謝の日", "1963-01-01": "元日", "1963-01-15": "成人の日", "1963-03-21": "春分の日", "1963-04-29": "天皇誕生日", "1963-05-03": "憲法記念日", "1963-05-05": "こどもの日", "1963-09-24": "秋分の日", "1963-11-03": "文化の日", "1963-11-23": "勤労感謝の日", "1964-01-01": "元日", "1964-01-15": "成人の日", "1964-03-20": "春分の日", "1964-04-29": "天皇誕生日", "1964-05-03": "憲法記念日", "1964-05-05": "こどもの日", "1964-09-23": "秋分の日", "1964-11-03": "文化の日", "1964-11-23": "勤労感謝の日", "1965-01-01": "元日", "1965-01-15": "成人の日", "1965-03-21": "春分の日", "1965-04-29": "天皇誕生日", "1965-05-03": "憲法記念日", "1965-05-05": "こどもの日", "1965-09-23": "秋分の日", "1965-11-03": "文化の日", "1965-11-23": "勤労感謝の日", "1966-01-01": "元日", "1966-01-15": "成人の日", "1966-03-21": "春分の日", "1966-04-29": "天皇誕生日", "1966-05-03": "憲法記念日", "1966-05-05": "こどもの日", "1966-09-15": "敬老の日", "1966-09-23": "秋分の日", "1966-10-10": "体育の日", "1966-11-03": "文化の日", "1966-11-23": "勤労感謝の日", "1967-01-01": "元日", "1967-01-15": "成人の日", "1967-02-11": "建国記念の日", "1967-03-21": "春分の日", "1967-04-29": "天皇誕生日", "1967-05-03": "憲法記念日", "1967-05-05": "こどもの日", "1967-09-15": "敬老の日", "1967-09-24": "秋分の日", "1967-10-10": "体育の日", "1967-11-03": "文化の日", "1967-11-23": "勤労感謝の日", "1968-01-01": "元日", "1968-01-15": "成人の日", "1968-02-11": "建国記念の日", "1968-03-20": "春分の日", "1968-04-29": "天皇誕生日", "1968-05-03": "憲法記念日", "1968-05-05": "こどもの日", "1968-09-15": "敬老の日", "1968-09-23": "秋分の日", "1968-10-10": "体育の日", "1968-11-03": "文化の日", "1968-11-23": "勤労感謝の日", "1969-01-01": "元日", "1969-01-15": "成人の日", "1969-02-11": "建国記念の日", "1969-03-21": "春分の日", "1969-04-29": "天皇誕生日", "1969-05-03": "憲法記念日", "1969-05-05": "こどもの日", "1969-09-15": "敬老の日", "1969-09-23": "秋分の日", "1969-10-10": "体育の日", "1969-11-03": "文化の日", "1969-11-23": "勤労感謝の日", "1970-01-01": "元日", "1970-01-15": "成人の日", "1970-02-11": "建国記念の日", "1970-03-21": "春分の日", "1970-04-29": "天皇誕生日", "1970-05-03": "憲法記念日", "1970-05-05": "こどもの日", "1970-09-15": "敬老の日", "1970-09-23": "秋分の日", "1970-10-10": "体育の日", "1970-11-03": "文化の日", "1970-11-23": "勤労感謝の日", "1971-01-01": "元日", "1971-01-15": "成人の日", "1971-02-11": "建国記念の日", "1971-03-21": "春分の日", "1971-04-29": "天皇誕生日", "1971-05-03": "憲法記念日", "1971-05-05": "こどもの日", "1971-09-15": "敬老の日", "1971-09-24": "秋分の日", "1971-10-10": "体育の日", "1971-11-03": "文化の日", "1971-11-23": "勤労感謝の日", "1972-01-01": "元日", "1972-01-15": "成人の日", "1972-02-11": "建国記念の日", "1972-03-20": "春分の日", "1972-04-29": "天皇誕生日", "1972-05-03": "憲法記念日", "1972-05-05": "こどもの日", "1972-09-15": "敬老の日", "1972-09-23": "秋分の日", "1972-10-10": "体育の日", "1972-11-03": "文化の日", "1972-11-23": "勤労感謝の日", "1973-01-01": "元日", "1973-01-15": "成人の日", "1973-02-11": "建国記念の日", "1973-03-21": "春分の日", "1973-04-29": "天皇誕生日", "1973-04-30": "休日", "1973-05-03": "憲法記念日", "1973-05-05": "こどもの日", "1973-09-15": "敬老の日", "1973-09-23": "秋分の日", "1973-09-24": "休日", "1973-10-10": "体育の日", "1973-11-03": "文化の日", "1973-11-23": "勤労感謝の日", "1974-01-01": "元日", "1974-01-15": "成人の日", "1974-02-11": "建国記念の日", "1974-03-21": "春分の日", "1974-04-29": "天皇誕生日", "1974-05-03": "憲法記念日", "1974-05-05": "こどもの日", "1974-05-06": "休日", "1974-09-15": "敬老の日", "1974-09-16": "休日", "1974-09-23": "秋分の日", "1974-10-10": "体育の日", "1974-11-03": "文化の日", "1974-11-04": "休日", "1974-11-23": "勤労感謝の日", "1975-01-01": "元日", "1975-01-15": "成人の日", "1975-02-11": "建国記念の日", "1975-03-21": "春分の日", "1975-04-29": "天皇誕生日", "1975-05-03": "憲法記念日", "1975-05-05": "こどもの日", "1975-09-15": "敬老の日", "1975-09-24": "秋分の日", "1975-10-10": "体育の日", "1975-11-03": "文化の日", "1975-11-23": "勤労感謝の日", "1975-11-24": "休日", "1976-01-01": "元日", "1976-01-15": "成人の日", "1976-02-11": "建国記念の日", "1976-03-20": "春分の日", "1976-04-29": "天皇誕生日", "1976-05-03": "憲法記念日", "1976-05-05": "こどもの日", "1976-09-15": "敬老の日", "1976-09-23": "秋分の日", "1976-10-10": "体育の日", "1976-10-11": "休日", "1976-11-03": "文化の日", "1976-11-23": "勤労感謝の日", "1977-01-01": "元日", "1977-01-15": "成人の日", "1977-02-11": "建国記念の日", "1977-03-21": "春分の日", "1977-04-29": "天皇誕生日", "1977-05-03": "憲法記念日", "1977-05-05": "こどもの日", "1977-09-15": "敬老の日", "1977-09-23": "秋分の日", "1977-10-10": "体育の日", "1977-11-03": "文化の日", "1977-11-23": "勤労感謝の日", "1978-01-01": "元日", "1978-01-02": "休日", "1978-01-15": "成人の日", "1978-01-16": "休日", "1978-02-11": "建国記念の日", "1978-03-21": "春分の日", "1978-04-29": "天皇誕生日", "1978-05-03": "憲法記念日", "1978-05-05": "こどもの日", "1978-09-15": "敬老の日", "1978-09-23": "秋分の日", "1978-10-10": "体育の日", "1978-11-03": "文化の日", "1978-11-23": "勤労感謝の日", "1979-01-01": "元日", "1979-01-15": "成人の日", "1979-02-11": "建国記念の日", "1979-02-12": "休日", "1979-03-21": "春分の日", "1979-04-29": "天皇誕生日", "1979-04-30": "休日", "1979-05-03": "憲法記念日", "1979-05-05": "こどもの日", "1979-09-15": "敬老の日", "1979-09-24": "秋分の日", "1979-10-10": "体育の日", "1979-11-03": "文化の日", "1979-11-23": "勤労感謝の日", "1980-01-01": "元日", "1980-01-15": "成人の日", "1980-02-11": "建国記念の日", "1980-03-20": "春分の日", "1980-04-29": "天皇誕生日", "1980-05-03": "憲法記念日", "1980-05-05": "こどもの日", "1980-09-15": "敬老の日", "1980-09-23": "秋分の日", "1980-10-10": "体育の日", "1980-11-03": "文化の日", "1980-11-23": "勤労感謝の日", "1980-11-24": "休日", "1981-01-01": "元日", "1981-01-15": "成人の日", "1981-02-11": "建国記念の日", "1981-03-21": "春分の日", "1981-04-29": "天皇誕生日", "1981-05-03": "憲法記念日", "1981-05-04": "休日", "1981-05-05": "こどもの日", "1981-09-15": "敬老の日", "1981-09-23": "秋分の日", "1981-10-10": "体育の日", "1981-11-03": "文化の日", "1981-11-23": "勤労感謝の日", "1982-01-01": "元日", "1982-01-15": "成人の日", "1982-02-11": "建国記念の日", "1982-03-21": "春分の日", "1982-03-22": "休日", "1982-04-29": "天皇誕生日", "1982-05-03": "憲法記念日", "1982-05-05": "こどもの日", "1982-09-15": "敬老の日", "1982-09-23": "秋分の日", "1982-10-10": "体育の日", "1982-10-11": "休日", "1982-11-03": "文化の日", "1982-11-23": "勤労感謝の日", "1983-01-01": "元日", "1983-01-15": "成人の日", "1983-02-11": "建国記念の日", "1983-03-21": "春分の日", "1983-04-29": "天皇誕生日", "1983-05-03": "憲法記念日", "1983-05-05": "こどもの日", "1983-09-15": "敬老の日", "1983-09-23": "秋分の日", "1983-10-10": "体育の日", "1983-11-03": "文化の日", "1983-11-23": "勤労感謝の日", "1984-01-01": "元日", "1984-01-02": "休日", "1984-01-15": "成人の日", "1984-01-16": "休日", "1984-02-11": "建国記念の日", "1984-03-20": "春分の日", "1984-04-29": "天皇誕生日", "1984-04-30": "休日", "1984-05-03": "憲法記念日", "1984-05-05": "こどもの日", "1984-09-15": "敬老の日", "1984-09-23": "秋分の日", "1984-09-24": "休日", "1984-10-10": "体育の日", "1984-11-03": "文化の日", "1984-11-23": "勤労感謝の日", "1985-01-01": "元日", "1985-01-15": "成人の日", "1985-02-11": "建国記念の日", "1985-03-21": "春分の日", "1985-04-29": "天皇誕生日", "1985-05-03": "憲法記念日", "1985-05-05": "こどもの日", "1985-05-06": "休日", "1985-09-15": "敬老の日", "1985-09-16": "休日", "1985-09-23": "秋分の日", "1985-10-10": "体育の日", "1985-11-03": "文化の日", "1985-11-04": "休日", "1985-11-23": "勤労感謝の日", "1986-01-01": "元日", "1986-01-15": "成人の日", "1986-02-11": "建国記念の日", "1986-03-21": "春分の日", "1986-04-29": "天皇誕生日", "1986-05-03": "憲法記念日", "1986-05-05": "こどもの日", "1986-09-15": "敬老の日", "1986-09-23": "秋分の日", "1986-10-10": "体育の日", "1986-11-03": "文化の日", "1986-11-23": "勤労感謝の日", "1986-11-24": "休日", "1987-01-01": "元日", "1987-01-15": "成人の日", "1987-02-11": "建国記念の日", "1987-03-21": "春分の日", "1987-04-29": "天皇誕生日", "1987-05-03": "憲法記念日", "1987-05-04": "休日", "1987-05-05": "こどもの日", "1987-09-15": "敬老の日", "1987-09-23": "秋分の日", "1987-10-10": "体育の日", "1987-11-03": "文化の日", "1987-11-23": "勤労感謝の日", "1988-01-01": "元日", "1988-01-15": "成人の日", "1988-02-11": "建国記念の日", "1988-03-20": "春分の日", "1988-03-21": "休日", "1988-04-29": "天皇誕生日", "1988-05-03": "憲法記念日", "1988-05-04": "休日", "1988-05-05": "こどもの日", "1988-09-15": "敬老の日", "1988-09-23": "秋分の日", "1988-10-10": "体育の日", "1988-11-03": "文化の日", "1988-11-23": "勤労感謝の日", "1989-01-01": "元日", "1989-01-02": "休日", "1989-01-15": "成人の日", "1989-01-16": "休日", "1989-02-11": "建国記念の日", "1989-02-24": "大喪の礼", "1989-03-21": "春分の日", "1989-04-29": "みどりの日", "1989-05-03": "憲法記念日", "1989-05-04": "休日", "1989-05-05": "こどもの日", "1989-09-15": "敬老の日", "1989-09-23": "秋分の日", "1989-10-10": "体育の日", "1989-11-03": "文化の日", "1989-11-23": "勤労感謝の日", "1989-12-23": "天皇誕生日", "1990-01-01": "元日", "1990-01-15": "成人の日", "1990-02-11": "建国記念の日", "1990-02-12": "休日", "1990-03-21": "春分の日", "1990-04-29": "みどりの日", "1990-04-30": "休日", "1990-05-03": "憲法記念日", "1990-05-04": "休日", "1990-05-05": "こどもの日", "1990-09-15": "敬老の日", "1990-09-23": "秋分の日", "1990-09-24": "休日", "1990-10-10": "体育の日", "1990-11-03": "文化の日", "1990-11-12": "即位礼正殿の儀", "1990-11-23": "勤労感謝の日", "1990-12-23": "天皇誕生日", "1990-12-24": "休日", "1991-01-01": "元日", "1991-01-15": "成人の日", "1991-02-11": "建国記念の日", "1991-03-21": "春分の日", "1991-04-29": "みどりの日", "1991-05-03": "憲法記念日", "1991-05-04": "休日", "1991-05-05": "こどもの日", "1991-05-06": "休日", "1991-09-15": "敬老の日", "1991-09-16": "休日", "1991-09-23": "秋分の日", "1991-10-10": "体育の日", "1991-11-03": "文化の日", "1991-11-04": "休日", "1991-11-23": "勤労感謝の日", "1991-12-23": "天皇誕生日", "1992-01-01": "元日", "1992-01-15": "成人の日", "1992-02-11": "建国記念の日", "1992-03-20": "春分の日", "1992-04-29": "みどりの日", "1992-05-03": "憲法記念日", "1992-05-04": "休日", "1992-05-05": "こどもの日", "1992-09-15": "敬老の日", "1992-09-23": "秋分の日", "1992-10-10": "体育の日", "1992-11-03": "文化の日", "1992-11-23": "勤労感謝の日", "1992-12-23": "天皇誕生日", "1993-01-01": "元日", "1993-01-15": "成人の日", "1993-02-11": "建国記念の日", "1993-03-20": "春分の日", "1993-04-29": "みどりの日", "1993-05-03": "憲法記念日", "1993-05-04": "休日", "1993-05-05": "こどもの日", "1993-06-09": "結婚の儀", "1993-09-15": "敬老の日", "1993-09-23": "秋分の日", "1993-10-10": "体育の日", "1993-10-11": "休日", "1993-11-03": "文化の日", "1993-11-23": "勤労感謝の日", "1993-12-23": "天皇誕生日", "1994-01-01": "元日", "1994-01-15": "成人の日", "1994-02-11": "建国記念の日", "1994-03-21": "春分の日", "1994-04-29": "みどりの日", "1994-05-03": "憲法記念日", "1994-05-04": "休日", "1994-05-05": "こどもの日", "1994-09-15": "敬老の日", "1994-09-23": "秋分の日", "1994-10-10": "体育の日", "1994-11-03": "文化の日", "1994-11-23": "勤労感謝の日", "1994-12-23": "天皇誕生日", "1995-01-01": "元日", "1995-01-02": "休日", "1995-01-15": "成人の日", "1995-01-16": "休日", "1995-02-11": "建国記念の日", "1995-03-21": "春分の日", "1995-04-29": "みどりの日", "1995-05-03": "憲法記念日", "1995-05-04": "休日", "1995-05-05": "こどもの日", "1995-09-15": "敬老の日", "1995-09-23": "秋分の日", "1995-10-10": "体育の日", "1995-11-03": "文化の日", "1995-11-23": "勤労感謝の日", "1995-12-23": "天皇誕生日", "1996-01-01": "元日", "1996-01-15": "成人の日", "1996-02-11": "建国記念の日", "1996-02-12": "休日", "1996-03-20": "春分の日", "1996-04-29": "みどりの日", "1996-05-03": "憲法記念日", "1996-05-04": "休日", "1996-05-05": "こどもの日", "1996-05-06": "休日", "1996-07-20": "海の日", "1996-09-15": "敬老の日", "1996-09-16": "休日", "1996-09-23": "秋分の日", "1996-10-10": "体育の日", "1996-11-03": "文化の日", "1996-11-04": "休日", "1996-11-23": "勤労感謝の日", "1996-12-23": "天皇誕生日", "1997-01-01": "元日", "1997-01-15": "成人の日", "1997-02-11": "建国記念の日", "1997-03-20": "春分の日", "1997-04-29": "みどりの日", "1997-05-03": "憲法記念日", "1997-05-05": "こどもの日", "1997-07-20": "海の日", "1997-07-21": "休日", "1997-09-15": "敬老の日", "1997-09-23": "秋分の日", "1997-10-10": "体育の日", "1997-11-03": "文化の日", "1997-11-23": "勤労感謝の日", "1997-11-24": "休日", "1997-12-23": "天皇誕生日", "1998-01-01": "元日", "1998-01-15": "成人の日", "1998-02-11": "建国記念の日", "1998-03-21": "春分の日", "1998-04-29": "みどりの日", "1998-05-03": "憲法記念日", "1998-05-04": "休日", "1998-05-05": "こどもの日", "1998-07-20": "海の日", "1998-09-15": "敬老の日", "1998-09-23": "秋分の日", "1998-10-10": "体育の日", "1998-11-03": "文化の日", "1998-11-23": "勤労感謝の日", "1998-12-23": "天皇誕生日", "1999-01-01": "元日", "1999-01-15": "成人の日", "1999-02-11": "建国記念の日", "1999-03-21": "春分の日", "1999-03-22": "休日", "1999-04-29": "みどりの日", "1999-05-03": "憲法記念日", "1999-05-04": "休日", "1999-05-05": "こどもの日", "1999-07-20": "海の日", "1999-09-15": "敬老の日", "1999-09-23": "秋分の日", "1999-10-10": "体育の日", "1999-10-11": "休日", "1999-11-03": "文化の日", "1999-11-23": "勤労感謝の日", "1999-12-23": "天皇誕生日", "2000-01-01": "元日", "2000-01-10": "成人の日", "2000-02-11": "建国記念の日", "2000-03-20": "春分の日", "2000-04-29": "みどりの日", "2000-05-03": "憲法記念日", "2000-05-04": "休日", "2000-05-05": "こどもの日", "2000-07-20": "海の日", "2000-09-15": "敬老の日", "2000-09-23": "秋分の日", "2000-10-09": "体育の日", "2000-11-03": "文化の日", "2000-11-23": "勤労感謝の日", "2000-12-23": "天皇誕生日", "2001-01-01": "元日", "2001-01-08": "成人の日", "2001-02-11": "建国記念の日", "2001-02-12": "休日", "2001-03-20": "春分の日", "2001-04-29": "みどりの日", "2001-04-30": "休日", "2001-05-03": "憲法記念日", "2001-05-04": "休日", "2001-05-05": "こどもの日", "2001-07-20": "海の日", "2001-09-15": "敬老の日", "2001-09-23": "秋分の日", "2001-09-24": "休日", "2001-10-08": "体育の日", "2001-11-03": "文化の日", "2001-11-23": "勤労感謝の日", "2001-12-23": "天皇誕生日", "2001-12-24": "休日", "2002-01-01": "元日", "2002-01-14": "成人の日", "2002-02-11": "建国記念の日", "2002-03-21": "春分の日", "2002-04-29": "みどりの日", "2002-05-03": "憲法記念日", "2002-05-04": "休日", "2002-05-05": "こどもの日", "2002-05-06": "休日", "2002-07-20": "海の日", "2002-09-15": "敬老の日", "2002-09-16": "休日", "2002-09-23": "秋分の日", "2002-10-14": "体育の日", "2002-11-03": "文化の日", "2002-11-04": "休日", "2002-11-23": "勤労感謝の日", "2002-12-23": "天皇誕生日", "2003-01-01": "元日", "2003-01-13": "成人の日", "2003-02-11": "建国記念の日", "2003-03-21": "春分の日", "2003-04-29": "みどりの日", "2003-05-03": "憲法記念日", "2003-05-05": "こどもの日", "2003-07-21": "海の日", "2003-09-15": "敬老の日", "2003-09-23": "秋分の日", "2003-10-13": "体育の日", "2003-11-03": "文化の日", "2003-11-23": "勤労感謝の日", "2003-11-24": "休日", "2003-12-23": "天皇誕生日", "2004-01-01": "元日", "2004-01-12": "成人の日", "2004-02-11": "建国記念の日", "2004-03-20": "春分の日", "2004-04-29": "みどりの日", "2004-05-03": "憲法記念日", "2004-05-04": "休日", "2004-05-05": "こどもの日", "2004-07-19": "海の日", "2004-09-20": "敬老の日", "2004-09-23": "秋分の日", "2004-10-11": "体育の日", "2004-11-03": "文化の日", "2004-11-23": "勤労感謝の日", "2004-12-23": "天皇誕生日", "2005-01-01": "元日", "2005-01-10": "成人の日", "2005-02-11": "建国記念の日", "2005-03-20": "春分の日", "2005-03-21": "休日", "2005-04-29": "みどりの日", "2005-05-03": "憲法記念日", "2005-05-04": "休日", "2005-05-05": "こどもの日", "2005-07-18": "海の日", "2005-09-19": "敬老の日", "2005-09-23": "秋分の日", "2005-10-10": "体育の日", "2005-11-03": "文化の日", "2005-11-23": "勤労感謝の日", "2005-12-23": "天皇誕生日", "2006-01-01": "元日", "2006-01-02": "休日", "2006-01-09": "成人の日", "2006-02-11": "建国記念の日", "2006-03-21": "春分の日", "2006-04-29": "みどりの日", "2006-05-03": "憲法記念日", "2006-05-04": "休日", "2006-05-05": "こどもの日", "2006-07-17": "海の日", "2006-09-18": "敬老の日", "2006-09-23": "秋分の日", "2006-10-09": "体育の日", "2006-11-03": "文化の日", "2006-11-23": "勤労感謝の日", "2006-12-23": "天皇誕生日", "2007-01-01": "元日", "2007-01-08": "成人の日", "2007-02-11": "建国記念の日", "2007-02-12": "休日", "2007-03-21": "春分の日", "2007-04-29": "昭和の日", "2007-04-30": "休日", "2007-05-03": "憲法記念日", "2007-05-04": "みどりの日", "2007-05-05": "こどもの日", "2007-07-16": "海の日", "2007-09-17": "敬老の日", "2007-09-23": "秋分の日", "2007-09-24": "休日", "2007-10-08": "体育の日", "2007-11-03": "文化の日", "2007-11-23": "勤労感謝の日", "2007-12-23": "天皇誕生日", "2007-12-24": "休日", "2008-01-01": "元日", "2008-01-14": "成人の日", "2008-02-11": "建国記念の日", "2008-03-20": "春分の日", "2008-04-29": "昭和の日", "2008-05-03": "憲法記念日", "2008-05-04": "みどりの日", "2008-05-05": "こどもの日", "2008-05-06": "休日", "2008-07-21": "海の日", "2008-09-15": "敬老の日", "2008-09-23": "秋分の日", "2008-10-13": "体育の日", "2008-11-03": "文化の日", "2008-11-23": "勤労感謝の日", "2008-11-24": "休日", "2008-12-23": "天皇誕生日", "2009-01-01": "元日", "2009-01-12": "成人の日", "2009-02-11": "建国記念の日", "2009-03-20": "春分の日", "2009-04-29": "昭和の日", "2009-05-03": "憲法記念日", "2009-05-04": "みどりの日", "2009-05-05": "こどもの日", "2009-05-06": "休日", "2009-07-20": "海の日", "2009-09-21": "敬老の日", "2009-09-22": "休日", "2009-09-23": "秋分の日", "2009-10-12": "体育の日", "2009-11-03": "文化の日", "2009-11-23": "勤労感謝の日", "2009-12-23": "天皇誕生日", "2010-01-01": "元日", "2010-01-11": "成人の日", "2010-02-11": "建国記念の日", "2010-03-21": "春分の日", "2010-03-22": "休日", "2010-04-29": "昭和の日", "2010-05-03": "憲法記念日", "2010-05-04": "みどりの日", "2010-05-05": "こどもの日", "2010-07-19": "海の日", "2010-09-20": "敬老の日", "2010-09-23": "秋分の日", "2010-10-11": "体育の日", "2010-11-03": "文化の日", "2010-11-23": "勤労感謝の日", "2010-12-23": "天皇誕生日", "2011-01-01": "元日", "2011-01-10": "成人の日", "2011-02-11": "建国記念の日", "2011-03-21": "春分の日", "2011-04-29": "昭和の日", "2011-05-03": "憲法記念日", "2011-05-04": "みどりの日", "2011-05-05": "こどもの日", "2011-07-18": "海の日", "2011-09-19": "敬老の日", "2011-09-23": "秋分の日", "2011-10-10": "体育の日", "2011-11-03": "文化の日", "2011-11-23": "勤労感謝の日", "2011-12-23": "天皇誕生日", "2012-01-01": "元日", "2012-01-02": "休日", "2012-01-09": "成人の日", "2012-02-11": "建国記念の日", "2012-03-20": "春分の日", "2012-04-29": "昭和の日", "2012-04-30": "休日", "2012-05-03": "憲法記念日", "2012-05-04": "みどりの日", "2012-05-05": "こどもの日", "2012-07-16": "海の日", "2012-09-17": "敬老の日", "2012-09-22": "秋分の日", "2012-10-08": "体育の日", "2012-11-03": "文化の日", "2012-11-23": "勤労感謝の日", "2012-12-23": "天皇誕生日", "2012-12-24": "休日", "2013-01-01": "元日", "2013-01-14": "成人の日", "2013-02-11": "建国記念の日", "2013-03-20": "春分の日", "2013-04-29": "昭和の日", "2013-05-03": "憲法記念日", "2013-05-04": "みどりの日", "2013-05-05": "こどもの日", "2013-05-06": "休日", "2013-07-15": "海の日", "2013-09-16": "敬老の日", "2013-09-23": "秋分の日", "2013-10-14": "体育の日", "2013-11-03": "文化の日", "2013-11-04": "休日", "2013-11-23": "勤労感謝の日", "2013-12-23": "天皇誕生日", "2014-01-01": "元日", "2014-01-13": "成人の日", "2014-02-11": "建国記念の日", "2014-03-21": "春分の日", "2014-04-29": "昭和の日", "2014-05-03": "憲法記念日", "2014-05-04": "みどりの日", "2014-05-05": "こどもの日", "2014-05-06": "休日", "2014-07-21": "海の日", "2014-09-15": "敬老の日", "2014-09-23": "秋分の日", "2014-10-13": "体育の日", "2014-11-03": "文化の日", "2014-11-23": "勤労感謝の日", "2014-11-24": "休日", "2014-12-23": "天皇誕生日", "2015-01-01": "元日", "2015-01-12": "成人の日", "2015-02-11": "建国記念の日", "2015-03-21": "春分の日", "2015-04-29": "昭和の日", "2015-05-03": "憲法記念日", "2015-05-04": "みどりの日", "2015-05-05": "こどもの日", "2015-05-06": "休日", "2015-07-20": "海の日", "2015-09-21": "敬老の日", "2015-09-22": "休日", "2015-09-23": "秋分の日", "2015-10-12": "体育の日", "2015-11-03": "文化の日", "2015-11-23": "勤労感謝の日", "2015-12-23": "天皇誕生日", "2016-01-01": "元日", "2016-01-11": "成人の日", "2016-02-11": "建国記念の日", "2016-03-20": "春分の日", "2016-03-21": "休日", "2016-04-29": "昭和の日", "2016-05-03": "憲法記念日", "2016-05-04": "みどりの日", "2016-05-05": "こどもの日", "2016-07-18": "海の日", "2016-08-11": "山の日", "2016-09-19": "敬老の日", "2016-09-22": "秋分の日", "2016-10-10": "体育の日", "2016-11-03": "文化の日", "2016-11-23": "勤労感謝の日", "2016-12-23": "天皇誕生日", "2017-01-01": "元日", "2017-01-02": "休日", "2017-01-09": "成人の日", "2017-02-11": "建国記念の日", "2017-03-20": "春分の日", "2017-04-29": "昭和の日", "2017-05-03": "憲法記念日", "2017-05-04": "みどりの日", "2017-05-05": "こどもの日", "2017-07-17": "海の日", "2017-08-11": "山の日", "2017-09-18": "敬老の日", "2017-09-23": "秋分の日", "2017-10-09": "体育の日", "2017-11-03": "文化の日", "2017-11-23": "勤労感謝の日", "2017-12-23": "天皇誕生日", "2018-01-01": "元日", "2018-01-08": "成人の日", "2018-02-11": "建国記念の日", "2018-02-12": "休日", "2018-03-21": "春分の日", "2018-04-29": "昭和の日", "2018-04-30": "休日", "2018-05-03": "憲法記念日", "2018-05-04": "みどりの日", "2018-05-05": "こどもの日", "2018-07-16": "海の日", "2018-08-11": "山の日", "2018-09-17": "敬老の日", "2018-09-23": "秋分の日", "2018-09-24": "休日", "2018-10-08": "体育の日", "2018-11-03": "文化の日", "2018-11-23": "勤労感謝の日", "2018-12-23": "天皇誕生日", "2018-12-24": "休日", "2019-01-01": "元日", "2019-01-14": "成人の日", "2019-02-11": "建国記念の日", "2019-03-21": "春分の日", "2019-04-29": "昭和の日", "2019-04-30": "休日", "2019-05-01": "休日（祝日扱い）", "2019-05-02": "休日", "2019-05-03": "憲法記念日", "2019-05-04": "みどりの日", "2019-05-05": "こどもの日", "2019-05-06": "休日", "2019-07-15": "海の日", "2019-08-11": "山の日", "2019-08-12": "休日", "2019-09-16": "敬老の日", "2019-09-23": "秋分の日", "2019-10-14": "体育の日（スポーツの日）", "2019-10-22": "休日（祝日扱い）", "2019-11-03": "文化の日", "2019-11-04": "休日", "2019-11-23": "勤労感謝の日", "2020-01-01": "元日", "2020-01-13": "成人の日", "2020-02-11": "建国記念の日", "2020-02-23": "天皇誕生日", "2020-02-24": "休日", "2020-03-20": "春分の日", "2020-04-29": "昭和の日", "2020-05-03": "憲法記念日", "2020-05-04": "みどりの日", "2020-05-05": "こどもの日", "2020-05-06": "休日", "2020-07-23": "海の日", "2020-07-24": "スポーツの日", "2020-08-10": "山の日", "2020-09-21": "敬老の日", "2020-09-22": "秋分の日", "2020-11-03": "文化の日", "2020-11-23": "勤労感謝の日", "2021-01-01": "元日", "2021-01-11": "成人の日", "2021-02-11": "建国記念の日", "2021-02-23": "天皇誕生日", "2021-03-20": "春分の日", "2021-04-29": "昭和の日", "2021-05-03": "憲法記念日", "2021-05-04": "みどりの日", "2021-05-05": "こどもの日", "2021-07-22": "海の日", "2021-07-23": "スポーツの日", "2021-08-08": "山の日", "2021-08-09": "休日", "2021-09-20": "敬老の日", "2021-09-23": "秋分の日", "2021-11-03": "文化の日", "2021-11-23": "勤労感謝の日", "2022-01-01": "元日", "2022-01-10": "成人の日", "2022-02-11": "建国記念の日", "2022-02-23": "天皇誕生日", "2022-03-21": "春分の日", "2022-04-29": "昭和の日", "2022-05-03": "憲法記念日", "2022-05-04": "みどりの日", "2022-05-05": "こどもの日", "2022-07-18": "海の日", "2022-08-11": "山の日", "2022-09-19": "敬老の日", "2022-09-23": "秋分の日", "2022-10-10": "スポーツの日", "2022-11-03": "文化の日", "2022-11-23": "勤労感謝の日", "2023-01-01": "元日", "2023-01-02": "休日", "2023-01-09": "成人の日", "2023-02-11": "建国記念の日", "2023-02-23": "天皇誕生日", "2023-03-21": "春分の日", "2023-04-29": "昭和の日", "2023-05-03": "憲法記念日", "2023-05-04": "みどりの日", "2023-05-05": "こどもの日", "2023-07-17": "海の日", "2023-08-11": "山の日", "2023-09-18": "敬老の日", "2023-09-23": "秋分の日", "2023-10-09": "スポーツの日", "2023-11-03": "文化の日", "2023-11-23": "勤労感謝の日", "2024-01-01": "元日", "2024-01-08": "成人の日", "2024-02-11": "建国記念の日", "2024-02-12": "休日", "2024-02-23": "天皇誕生日", "2024-03-20": "春分の日", "2024-04-29": "昭和の日", "2024-05-03": "憲法記念日", "2024-05-04": "みどりの日", "2024-05-05": "こどもの日", "2024-05-06": "休日", "2024-07-15": "海の日", "2024-08-11": "山の日", "2024-08-12": "休日", "2024-09-16": "敬老の日", "2024-09-22": "秋分の日", "2024-09-23": "休日", "2024-10-14": "スポーツの日", "2024-11-03": "文化の日", "2024-11-04": "休日", "2024-11-23": "勤労感謝の日", "2025-01-01": "元日", "2025-01-13": "成人の日", "2025-02-11": "建国記念の日", "2025-02-23": "天皇誕生日", "2025-02-24": "休日", "2025-03-20": "春分の日", "2025-04-29": "昭和の日", "2025-05-03": "憲法記念日", "2025-05-04": "みどりの日", "2025-05-05": "こどもの日", "2025-05-06": "休日", "2025-07-21": "海の日", "2025-08-11": "山の日", "2025-09-15": "敬老の日", "2025-09-23": "秋分の日", "2025-10-13": "スポーツの日", "2025-11-03": "文化の日", "2025-11-23": "勤労感謝の日", "2025-11-24": "休日", "2026-01-01": "元日", "2026-01-12": "成人の日", "2026-02-11": "建国記念の日", "2026-02-23": "天皇誕生日", "2026-03-20": "春分の日", "2026-04-29": "昭和の日", "2026-05-03": "憲法記念日", "2026-05-04": "みどりの日", "2026-05-05": "こどもの日", "2026-05-06": "休日", "2026-07-20": "海の日", "2026-08-11": "山の日", "2026-09-21": "敬老の日", "2026-09-22": "休日", "2026-09-23": "秋分の日", "2026-10-12": "スポーツの日", "2026-11-03": "文化の日", "2026-11-23": "勤労感謝の日", "2027-01-01": "元日", "2027-01-11": "成人の日", "2027-02-11": "建国記念の日", "2027-02-23": "天皇誕生日", "2027-03-21": "春分の日", "2027-03-22": "休日", "2027-04-29": "昭和の日", "2027-05-03": "憲法記念日", "2027-05-04": "みどりの日", "2027-05-05": "こどもの日", "2027-07-19": "海の日", "2027-08-11": "山の日", "2027-09-20": "敬老の日", "2027-09-23": "秋分の日", "2027-10-11": "スポーツの日", "2027-11-03": "文化の日", "2027-11-23": "勤労感謝の日" } };

  // ../calc-api/src/calendar.js
  var HOLIDAYS = holidays_default.holidays;
  var RANGE = { from: holidays_default.from, to: holidays_default.to };
  var DAY = 864e5;
  var WEEKDAYS_JA = ["日", "月", "火", "水", "木", "金", "土"];
  var WEEKDAYS_EN = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
  var MAX_SPAN_DAYS = 3700;
  var MAX_ADD = 1e3;
  var ERAS = [
    { name: "令和", en: "Reiwa", abbr: "R", start: "2019-05-01" },
    { name: "平成", en: "Heisei", abbr: "H", start: "1989-01-08" },
    { name: "昭和", en: "Showa", abbr: "S", start: "1926-12-25" },
    { name: "大正", en: "Taisho", abbr: "T", start: "1912-07-30" },
    { name: "明治", en: "Meiji", abbr: "M", start: "1873-01-01", firstYear: 1868 }
  ];
  var iso = (t) => new Date(t).toISOString().slice(0, 10);
  var toTime = (s) => Date.parse(`${s}T00:00:00Z`);
  function parseDate(input) {
    const m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(
      String(input != null ? input : "").normalize("NFKC").trim()
    );
    if (!m) return null;
    const s = `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    const t = toTime(s);
    return Number.isNaN(t) || iso(t) !== s ? null : s;
  }
  function dateField(body, key) {
    const v = body == null ? void 0 : body[key];
    if (v == null || v === "") throw new InputError(`"${key}" is required (YYYY-MM-DD)`, key);
    const d = typeof v === "string" ? parseDate(v) : null;
    if (!d) throw new InputError(`"${key}" must be a valid date (YYYY-MM-DD)`, key);
    return d;
  }
  function inRange(d, key) {
    if (d < RANGE.from || d > RANGE.to) {
      throw new InputError(
        `"${key}" must be between ${RANGE.from} and ${RANGE.to} (holiday data range)`,
        key
      );
    }
  }
  function options(body) {
    var _a, _b, _c;
    const closed = (_a = body == null ? void 0 : body.closedDates) != null ? _a : [];
    if (!Array.isArray(closed) || closed.length > 400) {
      throw new InputError('"closedDates" must be an array of up to 400 dates', "closedDates");
    }
    const closedSet = new Set(
      closed.map((c) => {
        const d = typeof c === "string" ? parseDate(c) : null;
        if (!d) throw new InputError('"closedDates" must contain dates (YYYY-MM-DD)', "closedDates");
        return d;
      })
    );
    const yearEnd = (_b = body == null ? void 0 : body.yearEndClosure) != null ? _b : false;
    if (typeof yearEnd !== "boolean")
      throw new InputError('"yearEndClosure" must be boolean', "yearEndClosure");
    const weekend = (_c = body == null ? void 0 : body.weekendDays) != null ? _c : ["sat", "sun"];
    if (!Array.isArray(weekend) || !weekend.every((w) => WEEKDAYS_EN.includes(w))) {
      throw new InputError('"weekendDays" must be an array like ["sat","sun"]', "weekendDays");
    }
    return { closedSet, yearEnd, weekend: new Set(weekend.map((w) => WEEKDAYS_EN.indexOf(w))) };
  }
  function isBusinessDay(d, o) {
    const dow = new Date(toTime(d)).getUTCDay();
    if (o.weekend.has(dow) || HOLIDAYS[d] || o.closedSet.has(d)) return false;
    if (o.yearEnd) {
      const md = d.slice(5);
      if (md >= "12-29" || md <= "01-03") return false;
    }
    return true;
  }
  function toWareki(d) {
    var _a;
    const era = ERAS.find((e) => d >= e.start);
    if (!era) return null;
    const [y, m, day] = d.split("-").map(Number);
    const n = y - ((_a = era.firstYear) != null ? _a : Number(era.start.slice(0, 4))) + 1;
    const yearText = n === 1 ? "元" : String(n);
    return {
      era: era.name,
      eraEn: era.en,
      year: n,
      text: `${era.name}${yearText}年${m}月${day}日`,
      short: `${era.abbr}${n}.${m}.${day}`
    };
  }
  function fromWareki(text2) {
    var _a;
    const s = String(text2 != null ? text2 : "").normalize("NFKC").replace(/\s+/g, "");
    const m = /^(令和|平成|昭和|大正|明治|[RHSTMrhstm])(元|\d{1,2})[年./-](\d{1,2})[月./-](\d{1,2})日?$/.exec(
      s
    );
    if (!m) return null;
    const key = m[1].toUpperCase();
    const era = ERAS.find((e) => e.name === m[1] || e.abbr === key);
    const n = m[2] === "元" ? 1 : Number(m[2]);
    const y = ((_a = era.firstYear) != null ? _a : Number(era.start.slice(0, 4))) + n - 1;
    const d = parseDate(`${y}-${m[3]}-${m[4]}`);
    if (!d) return null;
    const next = ERAS[ERAS.indexOf(era) - 1];
    const outOfRange = d < era.start || next != null && d >= next.start;
    return { date: d, era: era.name, year: n, outOfEraRange: outOfRange };
  }
  function describe(d, o) {
    var _a;
    const dow = new Date(toTime(d)).getUTCDay();
    return {
      date: d,
      weekday: WEEKDAYS_EN[dow],
      weekdayJa: WEEKDAYS_JA[dow],
      isHoliday: Boolean(HOLIDAYS[d]),
      holidayName: (_a = HOLIDAYS[d]) != null ? _a : null,
      isBusinessDay: isBusinessDay(d, o),
      wareki: toWareki(d)
    };
  }
  function calendarDay(body) {
    const d = dateField(body, "date");
    inRange(d, "date");
    return describe(d, options(body));
  }
  function calendarHolidays(body) {
    const year = body == null ? void 0 : body.year;
    const [from, to] = [Number(RANGE.from.slice(0, 4)), Number(RANGE.to.slice(0, 4))];
    if (!Number.isInteger(year) || year < from || year > to) {
      throw new InputError(`"year" must be an integer between ${from} and ${to}`, "year");
    }
    const holidays = Object.entries(HOLIDAYS).filter(([d]) => d.startsWith(`${year}-`)).map(([date, name]) => ({
      date,
      name,
      weekdayJa: WEEKDAYS_JA[new Date(toTime(date)).getUTCDay()]
    }));
    return { year, count: holidays.length, holidays, source: "内閣府「国民の祝日」" };
  }
  function addBusinessDays(body) {
    const d = dateField(body, "date");
    inRange(d, "date");
    const days = body == null ? void 0 : body.days;
    if (!Number.isInteger(days) || Math.abs(days) > MAX_ADD) {
      throw new InputError(`"days" must be an integer between -${MAX_ADD} and ${MAX_ADD}`, "days");
    }
    const o = options(body);
    const step = days < 0 ? -1 : 1;
    let t = toTime(d);
    let left = Math.abs(days);
    while (left > 0) {
      t += step * DAY;
      const cur = iso(t);
      if (cur < RANGE.from || cur > RANGE.to) {
        throw new InputError(
          `Result is outside the holiday data range (${RANGE.from} to ${RANGE.to})`,
          "days"
        );
      }
      if (isBusinessDay(cur, o)) left--;
    }
    return { from: d, days, result: describe(iso(t), o) };
  }
  function countBusinessDays(body) {
    const from = dateField(body, "from");
    const to = dateField(body, "to");
    inRange(from, "from");
    inRange(to, "to");
    if (to < from) throw new InputError('"to" must be on or after "from"', "to");
    const span = (toTime(to) - toTime(from)) / DAY + 1;
    if (span > MAX_SPAN_DAYS)
      throw new InputError(`Range must be ${MAX_SPAN_DAYS} days or less`, "to");
    const o = options(body);
    let business = 0;
    const holidays = [];
    for (let t = toTime(from); t <= toTime(to); t += DAY) {
      const cur = iso(t);
      if (isBusinessDay(cur, o)) business++;
      if (HOLIDAYS[cur]) holidays.push({ date: cur, name: HOLIDAYS[cur] });
    }
    return { from, to, calendarDays: span, businessDays: business, holidays };
  }

  // ../calc-api/src/invoice.js
  var DAY2 = 864e5;
  var MAX_TERMS = 100;
  var HOLIDAY_RULES = ["previous", "next", "none"];
  var toTime2 = (s) => Date.parse(`${s}T00:00:00Z`);
  var lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  var ymd = (y, m, d) => `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  var dayIn = (y, m, day) => day === "end" ? lastDay(y, m) : Math.min(day, lastDay(y, m));
  function parseTerms(text2) {
    const s = String(text2 != null ? text2 : "").normalize("NFKC").replace(/\s+/g, "");
    const close = /(?:(\d{1,2})日|(月末|末日|末))締/.exec(s);
    const pay = /(当月|翌月|翌々月|(\d{1,2})[かヵヶカケ]?月後)の?(?:(\d{1,2})日|(末日|末))/.exec(s);
    const offset = !pay ? null : pay[1] === "当月" ? 0 : pay[1] === "翌月" ? 1 : pay[1] === "翌々月" ? 2 : Number(pay[2]);
    return {
      closingDay: close ? close[1] ? Number(close[1]) : "end" : null,
      paymentMonthOffset: offset,
      paymentDay: pay ? pay[3] ? Number(pay[3]) : "end" : null,
      holidayRule: /前営業日/.test(s) ? "previous" : /(翌|後)営業日/.test(s) ? "next" : null
    };
  }
  function holidayRuleOf(v) {
    if (v === "" || v == null) return void 0;
    const s = String(v).normalize("NFKC").trim().toLowerCase();
    if (HOLIDAY_RULES.includes(s)) return s;
    if (/^前(営業日|日)?$/.test(s)) return "previous";
    if (/^(翌|後)(営業日|日)?$/.test(s)) return "next";
    if (/^(なし|そのまま|ずらさない)$/.test(s)) return "none";
    throw new InputError(`休日の扱いは「前」「翌」「なし」のどれか: ${v}`, "holidayRule");
  }
  function dayField(v, key) {
    if (v === "end") return v;
    if (!Number.isInteger(v) || v < 1 || v > 31) {
      throw new InputError(`"${key}" must be an integer 1-31 or "end"`, key);
    }
    return v;
  }
  function paymentDate(body) {
    const d = dateField(body, "date");
    inRange(d, "date");
    let parsed = {};
    if ((body == null ? void 0 : body.terms) != null) {
      if (typeof body.terms !== "string" || body.terms.length > MAX_TERMS) {
        throw new InputError(`"terms" must be a string up to ${MAX_TERMS} characters`, "terms");
      }
      parsed = parseTerms(body.terms);
    }
    const pick = (key, fallback) => {
      var _a, _b;
      return (_b = (_a = body == null ? void 0 : body[key]) != null ? _a : parsed[key]) != null ? _b : fallback;
    };
    const closingDay = dayField(pick("closingDay", "end"), "closingDay");
    const paymentDay = pick("paymentDay", null);
    if (paymentDay == null) {
      throw new InputError(
        (body == null ? void 0 : body.terms) != null ? '"terms" must include the payment day (e.g. "末締め翌月25日払い")' : '"paymentDay" is required (1-31 or "end"), or give "terms"',
        (body == null ? void 0 : body.terms) != null ? "terms" : "paymentDay"
      );
    }
    dayField(paymentDay, "paymentDay");
    const offset = pick("paymentMonthOffset", 1);
    if (!Number.isInteger(offset) || offset < 0 || offset > 12) {
      throw new InputError('"paymentMonthOffset" must be an integer 0-12', "paymentMonthOffset");
    }
    const holidayRule = pick("holidayRule", "previous");
    if (!HOLIDAY_RULES.includes(holidayRule)) {
      throw new InputError(`"holidayRule" must be one of ${HOLIDAY_RULES.join(", ")}`, "holidayRule");
    }
    const o = options(body);
    let [y, m] = d.split("-").map(Number);
    if (Number(d.slice(8)) > dayIn(y, m, closingDay)) [y, m] = m === 12 ? [y + 1, 1] : [y, m + 1];
    const closingDate = ymd(y, m, dayIn(y, m, closingDay));
    const pm = m - 1 + offset;
    const [py, pmm] = [y + Math.floor(pm / 12), pm % 12 + 1];
    const scheduled = ymd(py, pmm, dayIn(py, pmm, paymentDay));
    let t = toTime2(scheduled);
    const step = holidayRule === "next" ? 1 : -1;
    for (let i = 0; holidayRule !== "none" && !isBusinessDay(new Date(t).toISOString().slice(0, 10), o); i++) {
      if (i > 60) throw new InputError("No business day found near the payment date", "closedDates");
      t += step * DAY2;
    }
    const pay = new Date(t).toISOString().slice(0, 10);
    if (pay < RANGE.from || pay > RANGE.to) {
      throw new InputError(
        `Payment date is outside the holiday data range (${RANGE.from} to ${RANGE.to})`,
        "date"
      );
    }
    return {
      date: d,
      terms: { closingDay, paymentMonthOffset: offset, paymentDay, holidayRule },
      closingDate,
      scheduledDate: scheduled,
      paymentDate: describe(pay, o),
      adjusted: pay !== scheduled,
      daysUntilPayment: (t - toTime2(d)) / DAY2
    };
  }
  function withholdingTax(base) {
    return base <= 1e6 ? Math.floor(base * 1021 / 1e4) : Math.floor((base - 1e6) * 2042 / 1e4) + 102100;
  }
  var TAX_RATES = [0, 0.08, 0.1];
  function withholding(body) {
    var _a, _b;
    const amount = body == null ? void 0 : body.amount;
    if (!Number.isInteger(amount) || amount < 0 || amount > 1e11) {
      throw new InputError('"amount" must be an integer between 0 and 100000000000 (JPY)', "amount");
    }
    const includes = (_a = body == null ? void 0 : body.amountIncludesTax) != null ? _a : false;
    if (typeof includes !== "boolean") {
      throw new InputError('"amountIncludesTax" must be boolean', "amountIncludesTax");
    }
    const rate = (_b = body == null ? void 0 : body.taxRate) != null ? _b : 0.1;
    if (!TAX_RATES.includes(rate)) {
      throw new InputError('"taxRate" must be 0, 0.08 or 0.1', "taxRate");
    }
    const consumptionTax = includes ? null : Math.floor(amount * Math.round(rate * 100) / 100);
    const invoiceTotal = includes ? amount : amount + consumptionTax;
    const tax = withholdingTax(amount);
    return {
      amount,
      amountIncludesTax: includes,
      taxRate: rate,
      consumptionTax,
      invoiceTotal,
      withholdingBase: amount,
      withholdingTax: tax,
      netPayment: invoiceTotal - tax,
      assumptions: [
        {
          code: "general_fees",
          ja: "原稿料・講演料・デザイン料など、個人への一般の報酬・料金として計算(司法書士・外交員・ホステス等の特例は対象外)",
          en: "General fees paid to individuals (writing, lectures, design, etc.). Special rules (judicial scriveners, sales agents, hostesses, etc.) are not covered"
        },
        ...includes ? [
          {
            code: "tax_not_separated",
            ja: "消費税を区分していないため、税込の額に源泉徴収",
            en: "Consumption tax is not stated separately, so withholding applies to the tax-inclusive amount"
          }
        ] : []
      ]
    };
  }

  // src/functions.js
  var Realty2 = globalThis.RealtyParser;
  var DAY3 = 864e5;
  var EPOCH = Date.UTC(1899, 11, 30);
  var isBlank = (v) => v === "" || v == null;
  function mapCells(matrix, fn) {
    const rows = Array.isArray(matrix) ? matrix : [[matrix]];
    return rows.map(
      (row) => (Array.isArray(row) ? row : [row]).map((c) => isBlank(c) ? "" : fn(c))
    );
  }
  function mapCells2(matrix, other, fn) {
    var _a;
    const single = !Array.isArray(other) || other.length === 1 && other[0].length === 1;
    const one = Array.isArray(other) ? (_a = other[0]) == null ? void 0 : _a[0] : other;
    return mapCells(matrix, (c) => c).map(
      (row, i) => row.map((c, j) => {
        var _a2;
        return c === "" ? "" : fn(c, single ? one : (_a2 = other[i]) == null ? void 0 : _a2[j]);
      })
    );
  }
  var serialToIso = (n) => new Date(EPOCH + Math.round(n) * DAY3).toISOString().slice(0, 10);
  var isoToSerial = (s) => (Date.parse(`${s}T00:00:00Z`) - EPOCH) / DAY3;
  function toIso(v) {
    if (typeof v === "number") return serialToIso(v);
    const s = String(v).trim();
    const iso2 = parseDate(s);
    if (iso2) return iso2;
    const w = fromWareki(s);
    if (w) return w.date;
    throw new Error(`日付として読めません: ${s}`);
  }
  function num2(v, name) {
    if (typeof v === "number") return v;
    const n = Realty2.parseYen(String(v));
    if (n == null) throw new Error(`${name}を数値として読めません: ${v}`);
    return n;
  }
  function closedDates(range2) {
    if (isBlank(range2)) return [];
    const list = Array.isArray(range2) ? range2.flat() : [range2];
    return list.filter((c) => !isBlank(c)).map(toIso);
  }
  var yes = (v) => v === true || String(v).toUpperCase() === "TRUE";
  function takeHome(monthly, bonus, age, prefecture) {
    const body = { monthlySalary: num2(monthly, "月給") };
    if (!isBlank(bonus)) body.annualBonus = num2(bonus, "賞与");
    if (!isBlank(age)) body.age = Math.floor(Number(age));
    if (!isBlank(prefecture)) body.prefecture = String(prefecture);
    return calculateTakeHome(body);
  }
  var TAKEHOME_PARAMS = [
    { name: "bonus", description: "年間の賞与の合計(省略時は0)" },
    { name: "age", description: "年齢(40〜64歳は介護保険がかかる。省略時は30)" },
    { name: "prefecture", description: "勤務先の協会けんぽの都道府県(省略時は東京都)" }
  ];
  function salaryOf(text2, work) {
    return analyzeSalary({ salary: String(text2), workConditions: isBlank(work) ? "" : String(work) });
  }
  var FUNCTIONS = [
    {
      name: "TAKEHOME",
      category: "takehome",
      description: "月給・賞与から1年間の手取り額を計算します(2026年度の率・会社員・独身で扶養なし)。",
      example: "=JP.TAKEHOME(300000) → 2,876,160",
      params: [
        { name: "monthlySalary", description: "月給(額面・各種手当込み)。範囲も可", range: true },
        ...TAKEHOME_PARAMS.map((p) => __spreadProps(__spreadValues({}, p), { optional: true }))
      ],
      result: "matrix",
      fn: (monthly, bonus, age, prefecture) => mapCells(monthly, (m) => takeHome(m, bonus, age, prefecture).takeHomeAnnual)
    },
    {
      name: "TAKEHOME_DETAIL",
      category: "takehome",
      description: "手取りの内訳(社会保険料・所得税・住民税)を2列の表で返します(2026年度の率)。",
      example: "=JP.TAKEHOME_DETAIL(300000)",
      params: [
        { name: "monthlySalary", description: "月給(額面・各種手当込み)" },
        ...TAKEHOME_PARAMS.map((p) => __spreadProps(__spreadValues({}, p), { optional: true }))
      ],
      result: "matrix",
      fn: (monthly, bonus, age, prefecture) => {
        const r = takeHome(monthly, bonus, age, prefecture);
        const s = r.socialInsurance;
        return [
          ["額面(年)", r.grossAnnual],
          ["健康保険", s.healthInsurance],
          ["介護保険", s.nursingCareInsurance],
          ["子ども・子育て支援金", s.childSupportLevy],
          ["厚生年金", s.pension],
          ["雇用保険", s.employmentInsurance],
          ["社会保険料 計", s.total],
          ["所得税", r.incomeTax],
          ["住民税(概算)", r.residentTax],
          ["手取り(年)", r.takeHomeAnnual],
          ["手取り(月平均)", r.takeHomeMonthlyAverage]
        ];
      }
    },
    {
      name: "ANNUAL_INCOME",
      category: "salary",
      description: "求人の給与欄の文章から年収の目安を計算します(月給×(12+賞与の月数)、年俸があれば年俸)。",
      example: '=JP.ANNUAL_INCOME("月給25万円～＋賞与年2回（4.5ヶ月分）") → 4,125,000',
      params: [
        { name: "salaryText", description: "給与欄の文章。範囲も可", range: true },
        { name: "workText", description: "勤務時間・休日の文章(省略可)", optional: true },
        { name: "which", description: '"min"(下限・省略時)または "max"(上限)', optional: true }
      ],
      result: "matrix",
      fn: (text2, work, which) => mapCells(text2, (t) => {
        var _a;
        const r = salaryOf(t, work);
        if (!r.annual) return "";
        return (_a = which === "max" ? r.annual.max : r.annual.min) != null ? _a : "";
      })
    },
    {
      name: "HOURLY_EQUIVALENT",
      category: "salary",
      description: "求人の給与欄の文章から、時給に換算した額(下限)を計算します。時給の求人や読めない時は空。",
      example: '=JP.HOURLY_EQUIVALENT("月給30万円", "実働8時間 年間休日125日")',
      params: [
        { name: "salaryText", description: "給与欄の文章。範囲も可", range: true },
        {
          name: "workText",
          description: "勤務時間・休日の文章(無ければ1日8時間・年間休日120日と仮定)",
          optional: true
        }
      ],
      result: "matrix",
      fn: (text2, work) => mapCells(text2, (t) => {
        var _a, _b;
        return (_b = (_a = salaryOf(t, work).hourlyEquivalent) == null ? void 0 : _a.min) != null ? _b : "";
      })
    },
    {
      name: "FIXED_OVERTIME",
      category: "salary",
      description: "求人の給与欄から固定残業代(みなし残業代)の金額を取り出します。書かれていなければ空。",
      example: '=JP.FIXED_OVERTIME("月給30万円（固定残業代40時間分・5万円を含む）") → 50,000',
      params: [{ name: "salaryText", description: "給与欄の文章。範囲も可", range: true }],
      result: "matrix",
      fn: (text2) => mapCells(text2, (t) => {
        var _a, _b;
        return (_b = (_a = salaryOf(t).fixedOvertime) == null ? void 0 : _a.amount) != null ? _b : "";
      })
    },
    {
      name: "YEN",
      category: "money",
      description: "日本語の金額の書き方を数値にします(例: 1億2000万円 → 120000000)。",
      example: '=JP.YEN("1億2000万円") → 120,000,000',
      params: [{ name: "text", description: "金額の文字列。範囲も可", range: true }],
      result: "matrix",
      fn: (text2) => mapCells(text2, (t) => {
        if (typeof t === "number") return t;
        const n = Realty2.parseYen(String(t));
        if (n == null) throw new Error(`金額として読めません: ${t}`);
        return n;
      })
    },
    {
      name: "AREA_M2",
      category: "money",
      description: "面積の書き方を㎡にします(例: 10坪 → 33.06)。数値はそのまま㎡とみなします。",
      example: '=JP.AREA_M2("10坪") → 33.06',
      params: [{ name: "text", description: "面積の文字列。範囲も可", range: true }],
      result: "matrix",
      fn: (text2) => mapCells(text2, (t) => {
        if (typeof t === "number") return t;
        const n = Realty2.parseArea(String(t));
        if (n == null) throw new Error(`面積として読めません: ${t}`);
        return Math.round(n * 100) / 100;
      })
    },
    {
      name: "TSUBO_PRICE",
      category: "money",
      description: "坪単価(価格または賃料 ÷ 坪数)を計算します。「4,980万円」「70.12㎡」のような書き方も可。",
      example: '=JP.TSUBO_PRICE("4,980万円", "70.12㎡")',
      params: [
        { name: "price", description: "価格または賃料。範囲も可", range: true },
        { name: "area", description: "面積(数値は㎡)" }
      ],
      result: "matrix",
      fn: (price, area) => {
        const a = typeof area === "number" ? area : Realty2.parseArea(String(area != null ? area : ""));
        if (!a) throw new Error(`面積として読めません: ${area}`);
        return mapCells(price, (p) => Math.round(num2(p, "価格") / (a * 0.3025)));
      }
    },
    {
      name: "LOAN_PAYMENT",
      category: "money",
      description: "住宅ローンの毎月の返済額(元利均等)を計算します。",
      example: '=JP.LOAN_PAYMENT("3000万円", 1, 35)',
      params: [
        { name: "principal", description: "借入額(「3,000万円」のような書き方も可)。範囲も可", range: true },
        { name: "ratePercent", description: "年利(%・省略時は1)", optional: true },
        { name: "years", description: "返済期間(年・省略時は35)", optional: true }
      ],
      result: "matrix",
      fn: (principal, ratePercent, years) => {
        const rate = isBlank(ratePercent) ? 1 : Number(ratePercent);
        const y = isBlank(years) ? 35 : Number(years);
        return mapCells(
          principal,
          (p) => Math.round(Realty2.monthlyPayment(num2(p, "借入額"), rate / 100, y))
        );
      }
    },
    {
      name: "WAREKI",
      category: "calendar",
      description: "日付を和暦にします(例: 2026/9/24 → 令和8年9月24日)。",
      example: "=JP.WAREKI(A2) → 令和8年9月24日",
      params: [
        { name: "date", description: "日付。範囲も可", range: true },
        { name: "format", description: '"long"(令和8年9月24日・省略時)または "short"(R8.9.24)', optional: true }
      ],
      result: "matrix",
      fn: (date, format) => mapCells(date, (d) => {
        const w = toWareki(toIso(d));
        if (!w) throw new Error("1873年より前の日付は和暦にできません");
        return format === "short" ? w.short : w.text;
      })
    },
    {
      name: "FROM_WAREKI",
      category: "calendar",
      description: "和暦の文字列を日付(シリアル値)にします(例: 令和6年4月1日、R6.4.1、平成元年1月8日)。",
      example: '=JP.FROM_WAREKI("R6.4.1") → 2024/4/1(表示形式を日付に)',
      params: [{ name: "text", description: "和暦の文字列。範囲も可", range: true }],
      result: "matrix",
      fn: (text2) => mapCells(text2, (t) => {
        const r = fromWareki(String(t));
        if (!r) throw new Error(`和暦として読めません: ${t}`);
        return isoToSerial(r.date);
      })
    },
    {
      name: "IS_HOLIDAY",
      category: "calendar",
      description: "日本の祝日(振替休日・国民の休日を含む)かどうかを返します。",
      example: "=JP.IS_HOLIDAY(A2)",
      params: [{ name: "date", description: "日付。範囲も可", range: true }],
      result: "matrix",
      fn: (date) => mapCells(date, (d) => Boolean(HOLIDAYS[toIso(d)]))
    },
    {
      name: "HOLIDAY_NAME",
      category: "calendar",
      description: "祝日の名前を返します(祝日でなければ空)。",
      example: "=JP.HOLIDAY_NAME(A2) → 敬老の日",
      params: [{ name: "date", description: "日付。範囲も可", range: true }],
      result: "matrix",
      fn: (date) => mapCells(date, (d) => {
        var _a;
        return (_a = HOLIDAYS[toIso(d)]) != null ? _a : "";
      })
    },
    {
      name: "IS_BUSINESS_DAY",
      category: "calendar",
      description: "営業日(土日・祝日・指定した休業日以外)かどうかを返します。",
      example: "=JP.IS_BUSINESS_DAY(A2, TRUE)",
      params: [
        { name: "date", description: "日付。範囲も可", range: true },
        { name: "yearEndClosure", description: "TRUE なら12月29日〜1月3日を休みとして扱う", optional: true },
        { name: "closedDates", description: "独自の休業日の範囲(省略可)", optional: true, range: true }
      ],
      result: "matrix",
      fn: (date, yearEnd, closed) => {
        const c = closedDates(closed);
        return mapCells(
          date,
          (d) => calendarDay({ date: toIso(d), yearEndClosure: yes(yearEnd), closedDates: c }).isBusinessDay
        );
      }
    },
    {
      name: "WORKDAY",
      category: "calendar",
      description: "○営業日後(マイナスなら前)の日付を返します。日本の祝日に対応した WORKDAY です。",
      example: "=JP.WORKDAY(A2, 5) → 5営業日後(表示形式を日付に)",
      params: [
        { name: "date", description: "開始日(この日は数えない)。範囲も可", range: true },
        { name: "days", description: "営業日数(マイナスで前へ)" },
        { name: "yearEndClosure", description: "TRUE なら12月29日〜1月3日を休みとして扱う", optional: true },
        { name: "closedDates", description: "独自の休業日の範囲(省略可)", optional: true, range: true }
      ],
      result: "matrix",
      fn: (date, days, yearEnd, closed) => {
        const c = closedDates(closed);
        return mapCells(date, (d) => {
          const r = addBusinessDays({
            date: toIso(d),
            days: Math.trunc(Number(days)),
            yearEndClosure: yes(yearEnd),
            closedDates: c
          });
          return isoToSerial(r.result.date);
        });
      }
    },
    {
      name: "NETWORKDAYS",
      category: "calendar",
      description: "期間の営業日数を返します(開始日・終了日を含む)。日本の祝日に対応した NETWORKDAYS です。",
      example: "=JP.NETWORKDAYS(A2, B2)",
      params: [
        { name: "startDate", description: "開始日" },
        { name: "endDate", description: "終了日" },
        { name: "yearEndClosure", description: "TRUE なら12月29日〜1月3日を休みとして扱う", optional: true },
        { name: "closedDates", description: "独自の休業日の範囲(省略可)", optional: true, range: true }
      ],
      result: "scalar",
      fn: (start, end, yearEnd, closed) => countBusinessDays({
        from: toIso(start),
        to: toIso(end),
        yearEndClosure: yes(yearEnd),
        closedDates: closedDates(closed)
      }).businessDays
    },
    {
      name: "HOLIDAYS",
      category: "calendar",
      description: "その年の祝日の一覧(日付のシリアル値と祝日名の2列)を返します(内閣府の公式データ)。",
      example: "=JP.HOLIDAYS(2026)",
      params: [{ name: "year", description: "年(例: 2026)" }],
      result: "matrix",
      fn: (year) => calendarHolidays({ year: Math.floor(Number(year)) }).holidays.map((h) => [
        isoToSerial(h.date),
        h.name
      ])
    },
    {
      name: "PAYMENT_DATE",
      category: "invoice",
      description: "支払条件(例: 末締め翌月25日払い)から支払日(シリアル値)を返します。土日・祝日なら前営業日にします。",
      example: '=JP.PAYMENT_DATE(A2, "末締め翌月25日払い") → 25日が休日なら前営業日(表示形式を日付に)',
      params: [
        { name: "date", description: "取引日・請求日。範囲も可", range: true },
        {
          name: "terms",
          description: "支払条件の文章(例: 末締め翌月25日払い、20日締め翌々月末日支払)。行ごとに違う条件なら範囲も可",
          range: true
        },
        {
          name: "holidayRule",
          description: '休日の時: "前"(前営業日・省略時)、"翌"(翌営業日)、"なし"。文章に「翌営業日」とあればそれに従う',
          optional: true
        },
        { name: "closedDates", description: "独自の休業日の範囲(省略可)", optional: true, range: true }
      ],
      result: "matrix",
      fn: (date, terms, holidayRule, closed) => {
        const rule = holidayRuleOf(holidayRule);
        const c = closedDates(closed);
        return mapCells2(date, terms, (d, t) => {
          if (isBlank(t)) throw new Error("支払条件がありません");
          const body = { date: toIso(d), terms: String(t), closedDates: c };
          if (rule) body.holidayRule = rule;
          return isoToSerial(paymentDate(body).paymentDate.date);
        });
      }
    },
    {
      name: "WITHHOLDING",
      category: "invoice",
      description: "報酬・料金の源泉徴収税額を返します(100万円までは10.21%、超える部分は20.42%。1円未満切り捨て)。",
      example: "=JP.WITHHOLDING(100000) → 10,210",
      params: [
        { name: "amount", description: "報酬の額(税抜。「10万円」のような書き方も可)。範囲も可", range: true },
        { name: "includesTax", description: "TRUE なら、消費税を区分していない税込の額として扱う", optional: true }
      ],
      result: "matrix",
      fn: (amount, includesTax) => mapCells(
        amount,
        (a) => withholding({ amount: Math.round(num2(a, "報酬の額")), amountIncludesTax: yes(includesTax) }).withholdingTax
      )
    }
  ];

  // src/entry.js
  for (const f of FUNCTIONS) {
    CustomFunctions.associate(f.name, (...args) => {
      try {
        return f.fn(...args);
      } catch (e) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          e && e.message ? e.message : String(e)
        );
      }
    });
  }
})();
