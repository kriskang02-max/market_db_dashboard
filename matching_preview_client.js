/**
 * GitHub Pages 등 정적 호스팅용 매칭 미리보기 (matching_pptx.build_preview 이식).
 * PPTX 생성은 로컬 serve_dashboard.py API가 필요합니다.
 */
(function (global) {
  "use strict";

  var SECTOR_ROWS = [
    ["국고채", "국고"],
    ["특은채 AAA", "특은"],
    ["특수/시은채 AAA", "특수시은"],
    ["카드채 AA", "카드"],
    ["캐피탈채 AA-", "캐피탈"],
  ];
  var TENORS = ["1년", "1.5년", "2년", "2.5년", "3년"];
  var TENOR_FIELDS = {
    "6개월": "6월이하(당일)",
    "1년": "1년이하(당일)",
    "1.5년": "1.5년이하(당일)",
    "2년": "2년이하(당일)",
    "2.5년": "2.5년이하(당일)",
    "3년": "3년이하(당일)",
  };
  var PREV_TENOR = {
    "1년": "6개월",
    "1.5년": "1년",
    "2년": "1.5년",
    "2.5년": "2년",
    "3년": "2.5년",
  };
  var PREV_TENOR_1Y = {
    "1년": null,
    "1.5년": "6개월",
    "2년": "1년",
    "2.5년": "1.5년",
    "3년": "2년",
  };
  var INST_MAP = {
    국고: "시가평가 3사평균 국고채권",
    특은: "시가평가 3사평균 산업은행",
    한전: "시가평가 3사평균 한국전력공사",
    시은: "시가평가 3사평균 신한은행",
    카드: "시가평가 3사평균 신한카드",
    캐피탈: "시가평가 3사평균 신한캐피탈",
    CD: "CD 3사평균 AAA",
    REPO: "REPO 종합 1일물",
  };
  var COMBOS_100 = [
    { tenor: "1년", label: "특은 100% + 특은 80%", legs: [["특은", 1.0], ["특은", 0.8]], rwa: "0%" },
    { tenor: "1년", label: "특은 100% + 특수/시은 80%", legs: [["특은", 1.0], ["특수시은", 0.8]], rwa: "~20%" },
    { tenor: "1년", label: "특수/시은 100% + 여전 80%", legs: [["특수시은", 1.0], ["여전", 0.8]], rwa: "20~40%" },
    { tenor: "1.5년", label: "특은 100% + 특은 80%", legs: [["특은", 1.0], ["특은", 0.8]], rwa: "0%" },
    { tenor: "1.5년", label: "특은 100% + 특수/시은 80%", legs: [["특은", 1.0], ["특수시은", 0.8]], rwa: "~20%" },
    { tenor: "1.5년", label: "특수/시은 100% + 여전 80%", legs: [["특수시은", 1.0], ["여전", 0.8]], rwa: "20~40%" },
    { tenor: "2년", label: "특은 100% + 특은 80%", legs: [["특은", 1.0], ["특은", 0.8]], rwa: "0%" },
    { tenor: "2년", label: "특은 100% + 특수/시은 80%", legs: [["특은", 1.0], ["특수시은", 0.8]], rwa: "~20%" },
    { tenor: "2년", label: "특수/시은 100% + 여전 80%", legs: [["특수시은", 1.0], ["여전", 0.8]], rwa: "20~40%" },
  ];
  var COMBOS_200 = [
    { tenor: "1년", label: "특은 160% + 특은 100%", legs: [["특은", 1.6], ["특은", 1.0], ["CD", 0.4]], rwa: "0%" },
    { tenor: "1년", label: "특은 160% + 특수/시은 100%", legs: [["특은", 1.6], ["특수시은", 1.0], ["CD", 0.4]], rwa: "~20%" },
    { tenor: "1년", label: "특수/시은 160% + 여전 100%", legs: [["특수시은", 1.6], ["여전", 1.0], ["CD", 0.4]], rwa: "20~40%" },
    { tenor: "1.5년", label: "특은 160% + 특은 100%", legs: [["특은", 1.6], ["특은", 1.0], ["CD", 0.4]], rwa: "0%" },
    { tenor: "1.5년", label: "특은 160% + 특수/시은 100%", legs: [["특은", 1.6], ["특수시은", 1.0], ["CD", 0.4]], rwa: "~20%" },
    { tenor: "1.5년", label: "특수/시은 160% + 여전 100%", legs: [["특수시은", 1.6], ["여전", 1.0], ["CD", 0.4]], rwa: "20~40%" },
    { tenor: "2년", label: "특은 160% + 특은 100%", legs: [["특은", 1.6], ["특은", 1.0], ["CD", 0.4]], rwa: "0%" },
    { tenor: "2년", label: "특은 160% + 특수/시은 100%", legs: [["특은", 1.6], ["특수시은", 1.0], ["CD", 0.4]], rwa: "~20%" },
    { tenor: "2년", label: "특수/시은 160% + 여전 100%", legs: [["특수시은", 1.6], ["여전", 1.0], ["CD", 0.4]], rwa: "20~40%" },
  ];

  var COLOR_CHG_UP = "#DC2626";
  var COLOR_CHG_DOWN = "#2563EB";
  var COLOR_CHG_FLAT = "#64748B";
  var COLOR_LEVEL = "#1E293B";
  var COLOR_WHITE = "#FFFFFF";

  var _seriesCache = null;
  var _seriesPromise = null;

  function num(v) {
    if (v == null) return null;
    var s = String(v).trim().replace(/,/g, "");
    if (!s || s === "—" || s === "-" || s === "NA" || s === "nan") return null;
    var n = parseFloat(s);
    return Number.isFinite(n) ? n : null;
  }

  function parseYmd(s) {
    if (!s || !/^\d{4}-\d{2}-\d{2}/.test(s)) return null;
    var d = new Date(s.slice(0, 10) + "T00:00:00");
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function csvUrl() {
    try {
      return new URL("matching_data.csv", location.href).href;
    } catch (_) {
      return "matching_data.csv";
    }
  }

  function parseSeriesFromCsv(text) {
    var out = Object.create(null);
    var lines = text.split(/\r?\n/);
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i];
      if (!line) continue;
      var parts = line.split(",");
      if (parts.length < 4) continue;
      var date = (parts[0] || "").trim();
      var inst = (parts[1] || "").trim();
      var field = (parts[2] || "").trim();
      var val = num(parts.slice(3).join(","));
      if (!date || !inst || !field || val == null) continue;
      if (!out[inst]) out[inst] = Object.create(null);
      if (!out[inst][field]) out[inst][field] = Object.create(null);
      out[inst][field][date] = val;
    }
    return out;
  }

  function loadSeries() {
    if (_seriesCache) return Promise.resolve(_seriesCache);
    if (_seriesPromise) return _seriesPromise;
    _seriesPromise = fetch(csvUrl())
      .then(function (res) {
        if (!res.ok) throw new Error("matching_data.csv 로드 실패 (" + res.status + ")");
        return res.text();
      })
      .then(function (text) {
        _seriesCache = parseSeriesFromCsv(text);
        return _seriesCache;
      })
      .catch(function (err) {
        _seriesPromise = null;
        throw err;
      });
    return _seriesPromise;
  }

  function availableAsOfDates(series) {
    var dates = Object.create(null);
    var ktb = series[INST_MAP["국고"]] || {};
    Object.keys(TENOR_FIELDS).forEach(function (t) {
      var field = TENOR_FIELDS[t];
      var map = ktb[field] || {};
      Object.keys(map).forEach(function (d) {
        dates[d] = true;
      });
    });
    return Object.keys(dates).sort().reverse();
  }

  function pickPrevWeekDate(datesAsc, asof) {
    var target = parseYmd(asof);
    if (!target) return null;
    var want = new Date(target.getTime() - 7 * 86400000);
    var best = null;
    var bestDiff = null;
    for (var i = 0; i < datesAsc.length; i++) {
      var d = datesAsc[i];
      var dt = parseYmd(d);
      if (!dt || dt >= target) continue;
      var diff = Math.abs((dt - want) / 86400000);
      var bestDt = best ? parseYmd(best) : null;
      if (best == null || diff < bestDiff || (diff === bestDiff && dt > bestDt)) {
        best = d;
        bestDiff = diff;
      }
    }
    return best;
  }

  function sectorRate(series, sectorKey, tenor, date) {
    var field = TENOR_FIELDS[tenor];
    if (!field) return null;
    if (sectorKey === "특수시은") {
      var a = ((series[INST_MAP["한전"]] || {})[field] || {})[date];
      var b = ((series[INST_MAP["시은"]] || {})[field] || {})[date];
      if (a == null || b == null) return null;
      return (a + b) / 2;
    }
    if (sectorKey === "여전") {
      var ca = ((series[INST_MAP["카드"]] || {})[field] || {})[date];
      var cb = ((series[INST_MAP["캐피탈"]] || {})[field] || {})[date];
      if (ca == null && cb == null) return null;
      if (ca == null) return cb;
      if (cb == null) return ca;
      return (ca + cb) / 2;
    }
    var inst = INST_MAP[sectorKey];
    if (!inst) return null;
    return ((series[inst] || {})[field] || {})[date];
  }

  function curveSp(series, sectorKey, tenor, date) {
    var prev = PREV_TENOR[tenor];
    if (!prev) return null;
    var a = sectorRate(series, sectorKey, tenor, date);
    var b = sectorRate(series, sectorKey, prev, date);
    if (a == null || b == null) return null;
    return a - b;
  }

  function tenorYears(tenor) {
    var t = String(tenor || "").trim();
    if (t === "3개월") return 0.25;
    if (t === "6개월") return 0.5;
    if (t.slice(-1) === "년") {
      var n = parseFloat(t.slice(0, -1));
      return Number.isFinite(n) ? n : null;
    }
    return null;
  }

  function expectedReturn(series, sectorKey, tenor, date) {
    var ytm = sectorRate(series, sectorKey, tenor, date);
    var ty = tenorYears(tenor);
    if (ytm == null || ty == null) return null;
    var roll = ty - 1;
    if (roll <= 0) return ytm;
    var prevT = PREV_TENOR_1Y[tenor];
    if (!prevT) return ytm;
    var prevYtm = sectorRate(series, sectorKey, prevT, date);
    if (prevYtm == null) return null;
    return ytm + (ytm - prevYtm) * roll;
  }

  function creditSp(series, sectorKey, tenor, date) {
    if (sectorKey === "국고") return 0;
    var a = sectorRate(series, sectorKey, tenor, date);
    var b = sectorRate(series, "국고", tenor, date);
    if (a == null || b == null) return null;
    return (a - b) * 100;
  }

  function percentileRank(values, current) {
    if (!values || !values.length || current == null || !Number.isFinite(current)) return null;
    var clean = values.filter(function (v) {
      return v != null && Number.isFinite(v);
    });
    if (!clean.length) return null;
    var below = 0;
    var equal = 0;
    for (var i = 0; i < clean.length; i++) {
      if (clean[i] < current) below++;
      else if (clean[i] === current) equal++;
    }
    return ((below + 0.5 * equal) / clean.length) * 100;
  }

  function histMetric(series, sectorKey, tenor, asof, kind) {
    var end = parseYmd(asof) || new Date();
    var start = new Date(end.getTime() - 365.25 * 10 * 86400000);
    var field = TENOR_FIELDS[tenor];
    var keys = [];
    if (sectorKey === "특수시은") keys = [INST_MAP["한전"], INST_MAP["시은"]];
    else if (sectorKey === "여전") keys = [INST_MAP["카드"], INST_MAP["캐피탈"]];
    else keys = [INST_MAP[sectorKey] || ""];
    var dateSet = Object.create(null);
    keys.forEach(function (inst) {
      Object.keys(((series[inst] || {})[field] || {})).forEach(function (d) {
        dateSet[d] = true;
      });
    });
    var out = [];
    Object.keys(dateSet)
      .sort()
      .forEach(function (d) {
        var dt = parseYmd(d);
        if (!dt || dt < start || dt > end) return;
        var v;
        if (kind === "credit") v = creditSp(series, sectorKey, tenor, d);
        else {
          v = curveSp(series, sectorKey, tenor, d);
          if (v != null) v = v * 100;
        }
        if (v != null && Number.isFinite(v)) out.push(v);
      });
    return out;
  }

  function fmtPct(v, digits) {
    if (v == null || !Number.isFinite(v)) return "";
    return v.toFixed(digits == null ? 2 : digits) + "%";
  }

  function fmtChgBp(chgPct) {
    if (chgPct == null || !Number.isFinite(chgPct)) return "";
    var bp = Math.round(chgPct * 100);
    if (bp > 0) return "+" + bp + "bp";
    if (bp < 0) return bp + "bp";
    return "0bp";
  }

  function fmtCurveBp(bp) {
    if (bp == null || !Number.isFinite(bp)) return "";
    var n = Math.round(bp);
    if (n > 0) return "+" + n + "bp";
    if (n < 0) return n + "bp";
    return "0bp";
  }

  function fmtBpPct(bp, pct) {
    if (bp == null || !Number.isFinite(bp)) return "";
    var sign = bp > 0 ? "+" : "";
    var left = Math.abs(bp) >= 10 ? sign + bp.toFixed(0) + "bp" : sign + bp.toFixed(1) + "bp";
    if (pct == null || !Number.isFinite(pct)) return left;
    return left + "\n" + pct.toFixed(0) + "%";
  }

  function fmtCurveCell(bp, pct) {
    var left = fmtCurveBp(bp);
    if (!left) return "";
    if (pct == null || !Number.isFinite(pct)) return left;
    return left + "\n" + pct.toFixed(0) + "%";
  }

  function fmtMinpyeong(level, chg) {
    if (level == null) return "";
    if (chg == null) return fmtPct(level);
    return fmtPct(level) + "\n" + fmtChgBp(chg);
  }

  function chgTextColor(chgPct) {
    if (chgPct == null || !Number.isFinite(chgPct)) return COLOR_CHG_FLAT;
    var bp = Math.round(chgPct * 100);
    if (bp > 0) return COLOR_CHG_UP;
    if (bp < 0) return COLOR_CHG_DOWN;
    return COLOR_CHG_FLAT;
  }

  function lerpByte(a, b, t) {
    return Math.round(a + (b - a) * t);
  }

  function heatmapHex(pct) {
    if (pct == null || !Number.isFinite(pct)) return null;
    var t = Math.max(0, Math.min(100, pct)) / 100;
    var coral = [255, 241, 242];
    var mid = [255, 255, 255];
    var softLo = [239, 246, 255];
    var softHi = [191, 219, 254];
    var r, g, b;
    if (t <= 0.5) {
      var u = t * 2;
      r = lerpByte(coral[0], mid[0], u);
      g = lerpByte(coral[1], mid[1], u);
      b = lerpByte(coral[2], mid[2], u);
    } else {
      var u2 = (t - 0.5) * 2;
      if (u2 <= 0.4) {
        var uu = u2 / 0.4;
        r = lerpByte(mid[0], softLo[0], uu);
        g = lerpByte(mid[1], softLo[1], uu);
        b = lerpByte(mid[2], softLo[2], uu);
      } else {
        var uu2 = (u2 - 0.4) / 0.6;
        r = lerpByte(softLo[0], softHi[0], uu2);
        g = lerpByte(softLo[1], softHi[1], uu2);
        b = lerpByte(softLo[2], softHi[2], uu2);
      }
    }
    function hx(n) {
      return ("0" + n.toString(16).toUpperCase()).slice(-2);
    }
    return "#" + hx(r) + hx(g) + hx(b);
  }

  function buildGrids(series, asof) {
    var dateSet = Object.create(null);
    var ktb = series[INST_MAP["국고"]] || {};
    TENORS.forEach(function (t) {
      Object.keys(ktb[TENOR_FIELDS[t]] || {}).forEach(function (d) {
        dateSet[d] = true;
      });
    });
    var datesAsc = Object.keys(dateSet).sort();
    var prev = pickPrevWeekDate(datesAsc, asof);
    var minpyeong = Object.create(null);
    var expected = Object.create(null);
    var credit = Object.create(null);
    var curve = Object.create(null);
    SECTOR_ROWS.forEach(function (pair) {
      var key = pair[1];
      TENORS.forEach(function (tenor) {
        var cellKey = key + "|" + tenor;
        var cur = sectorRate(series, key, tenor, asof);
        var prv = prev ? sectorRate(series, key, tenor, prev) : null;
        var chg = cur != null && prv != null ? cur - prv : null;
        minpyeong[cellKey] = {
          level: cur,
          chg: chg,
          text: fmtMinpyeong(cur, chg),
          line_colors: [COLOR_LEVEL, chgTextColor(chg)],
          line_bolds: [true, true],
          fill: COLOR_WHITE,
        };
        var exp = expectedReturn(series, key, tenor, asof);
        var expPrv = prev ? expectedReturn(series, key, tenor, prev) : null;
        var expChg = exp != null && expPrv != null ? exp - expPrv : null;
        expected[cellKey] = {
          level: exp,
          chg: expChg,
          text: fmtMinpyeong(exp, expChg),
          line_colors: [COLOR_LEVEL, chgTextColor(expChg)],
          line_bolds: [true, true],
          fill: COLOR_WHITE,
        };
        var csp = creditSp(series, key, tenor, asof);
        var cspPct =
          csp != null ? percentileRank(histMetric(series, key, tenor, asof, "credit"), csp) : null;
        credit[cellKey] = {
          bp: csp,
          pct: cspPct,
          text: key !== "국고" ? fmtBpPct(csp, cspPct) : "",
          line_colors: key !== "국고" ? [COLOR_LEVEL, COLOR_LEVEL] : [null, null],
          line_bolds: [true, false],
          fill: key !== "국고" ? heatmapHex(cspPct) || COLOR_WHITE : COLOR_WHITE,
        };
        var cv = curveSp(series, key, tenor, asof);
        var cvBp = cv != null ? cv * 100 : null;
        var cvPct =
          cvBp != null ? percentileRank(histMetric(series, key, tenor, asof, "curve"), cvBp) : null;
        curve[cellKey] = {
          bp: cvBp,
          pct: cvPct,
          text: fmtCurveCell(cvBp, cvPct),
          line_colors: [COLOR_LEVEL, COLOR_LEVEL],
          line_bolds: [true, false],
          fill: cvPct != null ? heatmapHex(cvPct) || COLOR_WHITE : COLOR_WHITE,
        };
      });
    });
    return { asof: asof, prev_week: prev, minpyeong: minpyeong, expected: expected, credit: credit, curve: curve };
  }

  function cdRate(series, tenorLabel, date) {
    var fieldMap = {
      "1개월": "1월(적용일)",
      "2개월": "2월(적용일)",
      "3개월": "3월(적용일)",
      "4개월": "4월(적용일)",
      "6개월": "6월(적용일)",
      "9개월": "9월(적용일)",
      "1년": "1년(적용일)",
    };
    var field = fieldMap[tenorLabel];
    if (!field) return null;
    return ((series[INST_MAP["CD"]] || {})[field] || {})[date];
  }

  function matchingDuration(tenor, weightSum) {
    var ty = tenorYears(tenor);
    if (ty == null || weightSum == null) return null;
    return (ty - 0.5) * weightSum;
  }

  function portfolioBondDuration(rowsIn) {
    var total = 0;
    var any = false;
    (rowsIn || []).forEach(function (r) {
      var kind = String(r.kind || "").toLowerCase();
      var sector = String(r.sector || "").toUpperCase();
      if (kind === "repo" || sector === "REPO") return;
      var w = (num(r.weight) || 0) / 100;
      var d = matchingDuration(r.tenor || "", w);
      if (d == null) return;
      total += d;
      any = true;
    });
    return any ? total : null;
  }

  function comboReturns(grids, combos, scenarios, series, asof) {
    var minpyeong = grids.minpyeong;
    var rows = [];
    (combos || []).forEach(function (combo) {
      var tenor = combo.tenor;
      var asset = 0;
      var bondW = 0;
      var er = 0;
      var ok = true;
      (combo.legs || []).forEach(function (leg) {
        var sector = leg[0];
        var w = leg[1];
        if (sector === "CD") {
          var rate = cdRate(series || {}, "6개월", asof);
          if (rate == null) {
            ok = false;
            return;
          }
          asset += w;
          er += w * rate;
          return;
        }
        var key = sector === "여전" ? "캐피탈" : sector;
        var cell = minpyeong[key + "|" + tenor];
        if (!cell || cell.level == null) {
          ok = false;
          return;
        }
        asset += w;
        bondW += w;
        er += w * cell.level;
      });
      if (!ok) {
        rows.push({ tenor: tenor, label: combo.label, rwa: combo.rwa, vals: ["", "", ""], bep: "" });
        return;
      }
      var repoW = 1.0 - asset;
      var vals = [];
      var freezeRet = null;
      (scenarios || []).slice(0, 3).forEach(function (sc, i) {
        var repo = num(sc.avg_1y);
        if (repo == null) {
          vals.push("");
          return;
        }
        var ret = er + repoW * repo;
        vals.push(fmtPct(ret));
        if (i === 0) freezeRet = ret;
      });
      var bep = "";
      var dur = matchingDuration(tenor, bondW);
      if (freezeRet != null && dur && dur !== 0) bep = fmtPct(freezeRet / dur);
      rows.push({ tenor: tenor, label: combo.label, rwa: combo.rwa, vals: vals, bep: bep });
    });
    return rows;
  }

  function scenarioHeaderLabel(label) {
    var s = String(label || "").trim();
    if (!s) return "";
    if (s.indexOf("동결") >= 0) return "동결";
    var m = s.match(/(\d+)\s*회/);
    if (m) return m[1] + "회 인상";
    return s.split("(")[0].trim() || s;
  }

  function matchComboRowIndex(combos, portfolioRows) {
    var legs = [];
    var tenor = null;
    (portfolioRows || []).forEach(function (r) {
      var sector = r.sector || "";
      var kind = r.kind || "";
      if (sector === "REPO" || kind === "Repo" || sector === "CD" || kind === "CD") return;
      var key = sector === "여전" ? "캐피탈" : sector;
      var w = Math.round((num(r.weight) || 0) / 100 * 10000) / 10000;
      legs.push([key, w]);
      if (r.tenor) tenor = r.tenor;
    });
    if (!legs.length || !tenor) return null;
    legs.sort(function (a, b) {
      return a[0] === b[0] ? a[1] - b[1] : a[0] < b[0] ? -1 : 1;
    });
    for (var i = 0; i < combos.length; i++) {
      var combo = combos[i];
      if (combo.tenor !== tenor) continue;
      var cLegs = (combo.legs || [])
        .filter(function (leg) {
          return leg[0] !== "CD";
        })
        .map(function (leg) {
          return [leg[0] === "여전" ? "캐피탈" : leg[0], Math.round(Number(leg[1]) * 10000) / 10000];
        })
        .sort(function (a, b) {
          return a[0] === b[0] ? a[1] - b[1] : a[0] < b[0] ? -1 : 1;
        });
      if (JSON.stringify(cLegs) === JSON.stringify(legs)) return i;
    }
    return null;
  }

  function portfolioMetrics(series, grids, inputs) {
    var pf = inputs.portfolio || {};
    var rowsIn = pf.rows || [];
    var scenarios = inputs.repo_scenarios || [];
    var selIdx = parseInt(pf.repo_scenario_idx != null ? pf.repo_scenario_idx : 0, 10);
    if (!Number.isFinite(selIdx)) selIdx = 0;
    if (scenarios.length) selIdx = Math.max(0, Math.min(selIdx, scenarios.length - 1));
    else selIdx = 0;
    var selSc = scenarios[selIdx] || {};
    var selRepo = num(selSc.avg_1y);
    if (selRepo == null) selRepo = num(selSc.repo);
    var selLabel = scenarioHeaderLabel(selSc.label || "") || "동결";
    var asof = grids.asof;
    var detailRows = [];
    var contribSum = 0;
    var weightSum = 0;
    rowsIn.forEach(function (r) {
      var kind = r.kind || "";
      var name = r.name || "";
      var weight = num(r.weight) || 0;
      var rate = num(r.rate);
      var sector = r.sector || "";
      var tenor = r.tenor || "";
      if (rate == null) {
        if (kind === "Repo" || sector === "REPO") rate = selRepo;
        else if (kind === "CD" || sector === "CD") rate = cdRate(series, tenor || "6개월", asof);
        else {
          var key = sector === "여전" ? "캐피탈" : sector;
          var cell = grids.minpyeong[key + "|" + tenor];
          rate = cell ? cell.level : null;
        }
      }
      var contrib = rate != null ? (rate * weight) / 100 : null;
      if (contrib != null) contribSum += contrib;
      weightSum += weight;
      detailRows.push({
        kind: kind,
        name: name,
        sector: sector,
        tenor: tenor,
        rate: rate,
        weight: weight,
        contrib: contrib,
        rate_text: rate != null ? fmtPct(rate) : "[     ]%",
        weight_text: weight.toFixed(0) + "%",
        contrib_text: contrib != null ? fmtPct(contrib) : "[     ]%",
      });
    });
    var total = contribSum;
    var portDur = portfolioBondDuration(rowsIn);
    var bepPct = total != null && portDur ? total / portDur : null;
    var scenHeaders = [];
    var scenVals = [];
    scenarios.slice(0, 4).forEach(function (sc) {
      var short = scenarioHeaderLabel(sc.label || "");
      var fr = num(sc.final_rate);
      var avg = num(sc.avg_1y);
      if (avg == null) avg = num(sc.repo);
      scenHeaders.push(
        short + " (" + (fr != null ? fmtPct(fr) : "") + ")\nRepo " + (avg != null ? fmtPct(avg) : "")
      );
      var repo = num(sc.avg_1y);
      if (repo == null) {
        scenVals.push("");
        return;
      }
      var s = 0;
      var ok = true;
      for (var i = 0; i < rowsIn.length; i++) {
        var rr = rowsIn[i];
        var w = (num(rr.weight) || 0) / 100;
        var k = rr.kind || "";
        var sec = rr.sector || "";
        var ten = rr.tenor || "";
        if (k === "Repo" || sec === "REPO") s += w * repo;
        else if (k === "CD" || sec === "CD") {
          var cd = cdRate(series, ten || "6개월", asof);
          if (cd == null) {
            ok = false;
            break;
          }
          s += w * cd;
        } else {
          var ck = sec === "여전" ? "캐피탈" : sec;
          var c = grids.minpyeong[ck + "|" + ten];
          if (!c || c.level == null) {
            ok = false;
            break;
          }
          s += w * c.level;
        }
      }
      scenVals.push(ok ? fmtPct(s) : "");
    });
    var shocks = inputs.rate_shocks || [-0.6, -0.4, -0.2, 0.0, 0.2, 0.4, 0.6];
    var sens = shocks.map(function (sh) {
      if (total == null) return "";
      return fmtPct(total - sh * (portDur || 0));
    });
    return {
      detail_rows: detailRows,
      total: total,
      total_text: fmtPct(total),
      weight_sum: weightSum,
      duration: portDur,
      bep_pct: bepPct,
      bep_text: bepPct != null ? fmtPct(bepPct) : "[X.XX]%",
      return_text:
        total != null
          ? "예상수익률 " + fmtPct(total) + " (" + selLabel + " 기준)"
          : "예상수익률 [X.XX]% (" + selLabel + " 기준)",
      repo_scenario_idx: selIdx,
      repo_scenario_label: selLabel,
      scen_headers: scenHeaders,
      scen_vals: scenVals,
      sens_vals: sens,
      shocks: shocks,
    };
  }

  function buildPreview(series, inputs) {
    var dates = availableAsOfDates(series);
    var inp = inputs || {};
    var asof = String(inp.asof || "").trim() || (dates[0] || "");
    if (dates.length && dates.indexOf(asof) < 0) {
      var earlier = dates.filter(function (d) {
        return d <= asof;
      });
      asof = earlier[0] || dates[0];
    }
    inp.asof = asof;
    var grids = buildGrids(series, asof);
    var cond100 = comboReturns(grids, COMBOS_100, inp.repo_scenarios || [], series, asof);
    var cond200 = comboReturns(grids, COMBOS_200, inp.repo_scenarios || [], series, asof);
    var port = portfolioMetrics(series, grids, inp);
    var pfRows = (inp.portfolio || {}).rows || [];
    return {
      asof: asof,
      dates: dates.slice(0, 120),
      prev_week: grids.prev_week,
      minpyeong: grids.minpyeong,
      expected: grids.expected,
      credit: grids.credit,
      curve: grids.curve,
      cond100: cond100,
      cond200: cond200,
      cond100_hi: matchComboRowIndex(COMBOS_100, pfRows),
      cond200_hi: matchComboRowIndex(COMBOS_200, pfRows),
      portfolio: port,
      inputs: inp,
      sectors: SECTOR_ROWS.map(function (x) {
        return x[0];
      }),
      sector_keys: SECTOR_ROWS.map(function (x) {
        return x[1];
      }),
      tenors: TENORS.slice(),
      source: "client",
    };
  }

  function isStaticHost() {
    var h = (location.hostname || "").toLowerCase();
    return h.indexOf("github.io") >= 0 || h.indexOf("pages.dev") >= 0;
  }

  global.MatchingPreviewClient = {
    loadSeries: loadSeries,
    buildPreview: buildPreview,
    isStaticHost: isStaticHost,
    csvUrl: csvUrl,
  };
})(typeof window !== "undefined" ? window : globalThis);
