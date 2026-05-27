import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import localforage from "localforage";
import {
  Upload,
  Search,
  Package,
  MapPin,
  Warehouse,
  AlertCircle,
  CheckCircle2,
  Layers,
  TableProperties,
  Calculator,
  XCircle,
  Info,
  Settings,
  Truck,
  Scale,
  Maximize2,
  ArrowRight,
  Database,
  FileWarning,
  AlertTriangle,
  Lightbulb,
  GitMerge,
  FileDown,
  Rocket,
  CheckSquare,
  Globe,
  TrendingDown,
  ShieldAlert,
} from "lucide-react";
import * as XLSX from "xlsx";

// === 全局物流规则配置 (严格模式) ===
interface CarrierRule {
  maxSide: number;
  maxMid: number;
  maxMin: number;
  minSide: number;
  minMid: number;
  minMin: number;
  maxWeight: number;
  maxGirth: number;
  volRatio: number;
  maxVolume?: number;
}

const SHIPPING_RULES: Record<"DHL" | "GLS" | "DPD", CarrierRule> = {
  DHL: {
    maxSide: 200,
    maxMid: 60,
    maxMin: 60,
    minSide: 15,
    minMid: 11,
    minMin: 1,
    maxWeight: 31.5,
    maxGirth: 360,
    volRatio: 5000,
  },
  GLS: {
    maxSide: 150,
    maxMid: 80,
    maxMin: 60,
    minSide: 15,
    minMid: 11,
    minMin: 3,
    maxWeight: 40.0,
    maxGirth: 300,
    volRatio: 6000,
    maxVolume: 0.15,
  },
  DPD: {
    maxSide: 120,
    maxMid: 60,
    maxMin: 60,
    minSide: 17,
    minMid: 13,
    minMin: 3,
    maxWeight: 31.5,
    maxGirth: 300,
    volRatio: 5000,
    maxVolume: 0.15,
  },
};

const FREIGHT_RULE_SOURCE = "2026德朗司海外仓邮费规则DE-5.15更新";

const FREIGHT_FEES = {
  DHL: {
    domesticOversize: 168,
    internationalOversize: 176.4,
    peakSeason: 1.6,
    domesticReturn: 37.8,
    weightDiffPenalty: 18.5,
    falseWeightPenalty: 151.2,
    ukBrexit: 33.2,
    internationalReturnEu: 111.4,
    internationalReturnNonEu: 222.7,
  },
  GLS: {
    volume: 15.4,
    overLength: 16,
    nonConveyable: 8.6,
    extreme: 556,
    remoteIsland: 121.4,
    addressCorrection: 23.2,
  },
  DPD: {
    oversize: 33,
    volume: 33,
    overweight: 282.1,
    oversizeLevel2: 279.2,
    extreme: 488.9,
    remote: 141.1,
    addressCorrection: 5.2,
    pod: 112.9,
    tireHandling: 56.5,
    redelivery: 23.6,
    postcodeDomestic: 65.9,
    postcodeInternational: 112.9,
    labelIssue: 10,
    peakSeason: 2.1,
  },
};

// --- 精确安全的浮点数处理 ---
const exactRound = (num: number, decimals = 2) =>
  Number(Math.round(Number(num + "e" + decimals)) + "e-" + decimals);

const parseNumber = (value: unknown, fallback = 0) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const raw = String(value ?? "").trim();
  if (!raw) return fallback;

  const normalized = raw
    .replace(/\s/g, "")
    .replace(/[^\d,.-]/g, "")
    .replace(/(?!^)-/g, "");
  if (!normalized || normalized === "-" || normalized === "," || normalized === ".") {
    return fallback;
  }

  const lastComma = normalized.lastIndexOf(",");
  const lastDot = normalized.lastIndexOf(".");
  let numeric = normalized;

  if (lastComma !== -1 && lastDot !== -1) {
    const decimalSep = lastComma > lastDot ? "," : ".";
    const thousandSep = decimalSep === "," ? "." : ",";
    numeric = normalized
      .replace(new RegExp(`\\${thousandSep}`, "g"), "")
      .replace(decimalSep, ".");
  } else if (lastComma !== -1) {
    const commaParts = normalized.split(",");
    numeric =
      commaParts.length === 2 && commaParts[1].length <= 2
        ? normalized.replace(",", ".")
        : normalized.replace(/,/g, "");
  } else {
    const dotParts = normalized.split(".");
    numeric =
      dotParts.length > 2 || (dotParts.length === 2 && dotParts[1].length === 3)
        ? normalized.replace(/\./g, "")
        : normalized;
  }

  const parsed = Number(numeric);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const parsePositiveNumber = (value: unknown, fallback = 0) => {
  const parsed = parseNumber(value, fallback);
  return parsed > 0 ? parsed : fallback;
};

// --- 辅助函数：SKU等价前缀处理 (共享 EB / AM / T1 / T2 库存) ---
const SKU_PREFIXES = ["EB", "AM", "T1", "T2"] as const;

const getSkuBase = (sku: string) => {
  const upper = (sku || "").trim().toUpperCase();
  const matchedPrefix = SKU_PREFIXES.find((prefix) =>
    upper.startsWith(`${prefix}-`),
  );
  return matchedPrefix ? upper.substring(matchedPrefix.length + 1) : upper;
};

const getSkuWithPrefix = (
  sku: string,
  prefix: (typeof SKU_PREFIXES)[number],
) => {
  const base = getSkuBase(sku);
  return base ? `${prefix}-${base}` : "";
};

const getEquivalentSkus = (sku: string) => {
  const upper = (sku || "").trim().toUpperCase();
  const base = getSkuBase(upper);
  if (!base) return upper ? [upper] : [];
  return Array.from(
    new Set([upper, base, ...SKU_PREFIXES.map((prefix) => `${prefix}-${base}`)]),
  ).filter(Boolean);
};

const isGermanyDestination = (country: string) => {
  const normalized = (country || "").trim().toUpperCase();
  return ["DE", "DEU", "GERMANY", "GERMAN"].includes(normalized) ||
    normalized.includes("德国") ||
    normalized.includes("德國");
};

// --- 渠道发货尺寸重量限制拦截判定 ---
const isChannelShippable = (
  channelName: string,
  country: string,
  length: number,
  width: number,
  height: number,
  weight: number,
): boolean => {
  const upperChannel = channelName.toUpperCase();

  // 如果源数据缺乏严格尺寸/重量，暂不强制拦截，只对有明确长宽高的产品起效
  if (length <= 0 || width <= 0 || height <= 0 || weight <= 0) return true;

  const dims = [length, width, height].sort((a, b) => b - a);
  const maxL = dims[0],
    midW = dims[1],
    minH = dims[2];
  const girth = exactRound(maxL + 2 * (midW + minH), 2);

  if (upperChannel.includes("DPD")) {
    // 5.15规则：最长边>250、围长>330、重量>40进入超限高风险，系统直接拦截止损。
    if (maxL > 250) return false;
    if (girth > 330) return false;
    if (weight > 40) return false;
  }

  if (upperChannel.includes("DHL") && !isGermanyDestination(country)) {
    // 德国发境外DHL：最长边<150cm，长+2*(宽+高)≤300cm，常规上限30kg。
    if (maxL >= 150) return false;
    if (girth > 300) return false;
    if (weight > 30) return false;
  }

  return true;
};

// --- 子组件：规则项展示 ---
interface RuleItemProps {
  label: string;
  value: number | string;
  limit: number;
  unit: string;
  valid: boolean;
  isMin?: boolean;
}

const RuleItem: React.FC<RuleItemProps> = React.memo(
  ({ label, value, limit, unit, valid, isMin }) => (
    <div
      className={`flex items-center justify-between p-2 rounded border-b last:border-b-0 transition-all ${valid ? "border-slate-100 hover:bg-slate-50/50" : "border-rose-100 bg-rose-50/30"}`}
    >
      <span
        className={`text-[10px] font-black uppercase tracking-tighter truncate mr-2 ${valid ? "text-slate-400" : "text-rose-400"}`}
        title={label}
      >
        {label}
      </span>
      <div className="flex items-center gap-1.5 sm:gap-2 flex-shrink-0">
        <span
          className={`font-mono font-bold text-xs ${valid ? "text-slate-700" : "text-rose-600"}`}
        >
          {value}{" "}
          <span className="text-[9px] font-normal opacity-40">{unit}</span>
        </span>
        <div className={`w-px h-3 ${valid ? "bg-slate-200" : "bg-rose-200"}`} />
        <span
          className={`text-[9px] font-mono font-bold min-w-[40px] sm:min-w-[45px] text-right ${valid ? "text-slate-400" : "text-rose-500"}`}
        >
          {isMin ? "MIN" : "MAX"}:{limit}
        </span>
      </div>
    </div>
  ),
);

RuleItem.displayName = "RuleItem";

/* =========================================================================
   Module 3: Orders Optimizer 
========================================================================= */
interface OrdersOptimizerProps {
  tariffMeta: any;
  tariffData: any[][] | null;
  productInfoMap: Record<
    string,
    { l: number; w: number; h: number; wt: number; qty?: number }
  >;
}

const OrdersOptimizer: React.FC<OrdersOptimizerProps> = ({
  tariffMeta,
  tariffData,
  productInfoMap,
}) => {
  const [fileData, setFileData] = useState<{
    name: string;
    buffer: ArrayBuffer;
  } | null>(null);
  const [logs, setLogs] = useState<
    { id: number; type: "info" | "success" | "error"; msg: string }[]
  >([]);
  const [loading, setLoading] = useState(false);
  const [processedWb, setProcessedWb] = useState<XLSX.WorkBook | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addLog = (msg: string, type: "info" | "success" | "error" = "info") => {
    setLogs((prev) => [...prev, { id: Date.now() + Math.random(), type, msg }]);
  };

  const handleUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const buffer = event.target?.result as ArrayBuffer;
      setFileData({ name: file.name, buffer });
      setProcessedWb(null);
      setLogs([]);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const processExcel = async () => {
    if (!fileData) return;
    setLoading(true);
    setLogs([]);
    addLog("开始加载并解析数据表矩阵...", "info");

    setTimeout(() => {
      try {
        const wb = XLSX.read(fileData.buffer, { type: "array" });
        const sheetNames = wb.SheetNames;

        // --- Sheet 1 预留 ---
        if (sheetNames.length >= 1) {
          addLog(
            `挂载 Sheet 1: [${sheetNames[0]}] - 通用SKU优先处理 (规则预留)`,
            "success",
          );
        }

        const skuStockMap: Record<
          string,
          { wh: string; ageNum: number; avail: number }[]
        > = {};

        // --- Sheet 2: 库龄优先发货逻辑 ---
        if (sheetNames.length >= 2) {
          addLog(
            `挂载 Sheet 2: [${sheetNames[1]}] - 解析库龄情况，重置发货提醒...`,
            "info",
          );
          const ws2 = wb.Sheets[sheetNames[1]];
          const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1 }) as any[][];

          if (data2.length > 0) {
            const header2 = data2[0] || [];
            const findIdx = (kws: string[]) =>
              header2.findIndex((c) =>
                kws.some((k) => String(c).toUpperCase().includes(k)),
              );
            const skuIdx = findIdx(["SKU", "编码", "商品编码"]);
            const ageIdx = findIdx(["库龄", "AGE", "情况", "状态"]);
            const availIdx = findIdx(["可用", "AVAILABLE", "库存"]);
            const pendingIdx = findIdx(["待发", "PENDING", "待出"]);
            const whIdx = findIdx(["仓库", "仓", "WAREHOUSE"]);
            let priorityCol = 3; // D列 (0-indexed 3)

            let dColUpdated = 0;
            for (let i = 1; i < data2.length; i++) {
              const row = data2[i];
              if (!row || row.length === 0) continue;

              const sku = String(
                row[skuIdx] !== undefined ? row[skuIdx] : "",
              )
                .trim()
                .toUpperCase();
              const wh = String(
                row[whIdx] !== undefined ? row[whIdx] : "",
              ).trim();
              const ageStr = String(
                row[ageIdx] !== undefined ? row[ageIdx] : "",
              );

              const availStr = String(
                row[availIdx] !== undefined ? row[availIdx] : "0",
              ).replace(/,/g, "");
              const pendingStr = String(
                row[pendingIdx] !== undefined ? row[pendingIdx] : "0",
              ).replace(/,/g, "");
              const avail = parseNumber(availStr);
              const pending = parseNumber(pendingStr);

              let ageNum = 0;
              const nums = ageStr.match(/\d+/g);
              if (nums) ageNum = Math.max(...nums.map(Number));
              else if (ageStr.includes("长") || ageStr.includes("久"))
                ageNum = 999;

              if (sku && wh) {
                if (!skuStockMap[sku]) skuStockMap[sku] = [];
                skuStockMap[sku].push({ wh, ageNum, avail });
              }

              // 核心规则判断：长库龄并且有可用库存时
              const isLongAge = ageStr.includes("长") || ageNum >= 60; // 此处用60天泛指"长",也可只通过字符串'长'判断
              if (isLongAge && avail > 0) {
                if (pending === 0) {
                  row[priorityCol] = "否";
                  dColUpdated++;
                } else if (pending > 0) {
                  row[priorityCol] = "是";
                  dColUpdated++;
                }
              }
            }
            wb.Sheets[sheetNames[1]] = XLSX.utils.aoa_to_sheet(data2);
            addLog(
              `完成 Sheet 2 [库龄核对]: 共 ${dColUpdated} 条SKU命中优先发货状态 (D列) 更新。`,
              "success",
            );
          }
        } else {
          addLog(
            "未检测到 Sheet 2 (库龄及可用库存表)，已跳过库存拦截计算。",
            "error",
          );
        }

        // --- Sheet 3: 欧盟订单智能切仓 ---
        if (sheetNames.length >= 3) {
          addLog(
            `挂载 Sheet 3: [${sheetNames[2]}] - 执行防错换仓与智能资费切仓链路...`,
            "info",
          );
          const ws3 = wb.Sheets[sheetNames[2]];
          const data3 = XLSX.utils.sheet_to_json(ws3, { header: 1 }) as any[][];

          const TARGET_WH_LIST = ["粮仓", "粮2仓", "T2仓"];
          const PREFERRED_WH_LIST = [
            "谷仓",
            "易达云仓",
            "捎客海外仓",
            "大麦仓",
          ];

          let sheet3Updated = 0;
          let costCalcCount = 0;

          const isSameWarehouse = (sourceWh: string, targetWh: string) => {
            const source = String(sourceWh || "").trim();
            const target = String(targetWh || "").trim();
            return (
              Boolean(source && target) &&
              (source.includes(target) || target.includes(source))
            );
          };

          const getCombinedInventory = (tSku: string) => {
            const eqSkus = getEquivalentSkus(tSku);
            const whMap: Record<
              string,
              { wh: string; ageNum: number; avail: number }
            > = {};
            eqSkus.forEach((s) => {
              if (skuStockMap[s]) {
                skuStockMap[s].forEach((item) => {
                  if (!whMap[item.wh]) whMap[item.wh] = { ...item };
                  else {
                    whMap[item.wh].avail += item.avail;
                    whMap[item.wh].ageNum = Math.max(
                      whMap[item.wh].ageNum,
                      item.ageNum,
                    );
                  }
                });
              }
            });
            return Object.values(whMap);
          };

          const getSuggestedStockSku = (tSku: string, targetWh: string) => {
            const eqSkus = getEquivalentSkus(tSku);
            const stockBySku: Record<string, number> = {};

            eqSkus.forEach((s) => {
              skuStockMap[s]?.forEach((item) => {
                if (item.avail > 0 && isSameWarehouse(item.wh, targetWh)) {
                  stockBySku[s] = (stockBySku[s] || 0) + item.avail;
                }
              });
            });

            const preferredSkuOrder = [
              tSku.trim().toUpperCase(),
              getSkuWithPrefix(tSku, "AM"),
              getSkuWithPrefix(tSku, "EB"),
              getSkuWithPrefix(tSku, "T1"),
              getSkuWithPrefix(tSku, "T2"),
              getSkuBase(tSku),
              ...eqSkus,
            ];

            return (
              Array.from(new Set(preferredSkuOrder)).find(
                (s) => (stockBySku[s] || 0) > 0,
              ) || ""
            );
          };

          if (data3.length > 0) {
            const header3 = data3[0] || [];
            const skuIdx3 = header3.findIndex((c) =>
              String(c).toUpperCase().includes("SKU"),
            );
            const countryIdx3 = header3.findIndex(
              (c) =>
                String(c).includes("国家") ||
                String(c).toUpperCase().includes("COUNTRY") ||
                String(c).includes("目的国"),
            );
            // 寻找账号列以判定"欧盟账号"
            const accountIdx = header3.findIndex(
              (c) =>
                String(c).includes("账号") ||
                String(c).includes("店铺") ||
                String(c).includes("平台"),
            );

            const colD = 3; // D列填：是
            const colE = 4; // E列填：推荐仓库
            let colF = header3.findIndex(
              (c) => String(c).includes("仓库") || String(c).includes("发运仓"),
            );
            if (colF === -1) colF = 5; // 如果没找到明确标头，假设仓库在F列 (index 5)

            let cheapestWhCol = header3.indexOf("系统比价:最佳仓");
            if (cheapestWhCol === -1) {
              cheapestWhCol = header3.length;
              header3[cheapestWhCol] = "系统比价:最佳仓";
            }

            let cheapestChCol = header3.indexOf("系统比价:最佳渠道");
            if (cheapestChCol === -1) {
              cheapestChCol = header3.length;
              header3[cheapestChCol] = "系统比价:最佳渠道";
            }

            let costColIdx = header3.indexOf("系统比价:最低运费");
            if (costColIdx === -1) {
              costColIdx = header3.length;
              header3[costColIdx] = "系统比价:最低运费";
            }

            let stockStatusCol = header3.indexOf("比价仓库存情况");
            if (stockStatusCol === -1) {
              stockStatusCol = header3.length;
              header3[stockStatusCol] = "比价仓库存情况";
            }

            let suggestedSkuCol = header3.indexOf("建议发货SKU");
            if (suggestedSkuCol === -1) {
              suggestedSkuCol = header3.length;
              header3[suggestedSkuCol] = "建议发货SKU";
            }

            for (let i = 1; i < data3.length; i++) {
              const row = data3[i];
              if (!row || row.length < 2) continue;

              const sku = String(
                row[skuIdx3] !== undefined ? row[skuIdx3] : "",
              ).trim();
              const country =
                countryIdx3 !== -1
                  ? String(
                      row[countryIdx3] !== undefined ? row[countryIdx3] : "",
                    ).trim()
                  : "";
              const currentWh = String(
                row[colF] !== undefined ? row[colF] : "",
              ).trim();

              // 1. 欧盟账号防错拦截逻辑 (原需求)
              const isTargetWh = TARGET_WH_LIST.some((tw) =>
                currentWh.includes(tw),
              );
              if (isTargetWh) {
                row[colD] = "是";
                let suggestion = "/";

                const combinedStock = getCombinedInventory(sku);
                if (combinedStock.length > 0) {
                  const candidates = combinedStock.filter(
                    (s) =>
                      s.avail > 0 &&
                      PREFERRED_WH_LIST.some((pw) => s.wh.includes(pw)),
                  );
                  if (candidates.length > 0) {
                    candidates.sort((a, b) => b.ageNum - a.ageNum);
                    suggestion = candidates[0].wh;
                  }
                }
                row[colE] = suggestion;
                sheet3Updated++;
              }

              // 2. 「资费解析」最低运费/最优仓核算 (新需求)
              let bestPriceWH = "";
              let bestPriceChannel = "";
              let lowestCost = Infinity;

              if (tariffData && sku && country) {
                const {
                  skuIndex: tSkuIdx,
                  countryIndex: tCountryIdx,
                  warehouses: tWhs,
                  lengthIdx,
                  widthIdx,
                  heightIdx,
                  weightIdx,
                } = tariffMeta;

                const equivalentOrderSkus = getEquivalentSkus(sku);
                const tariffRow = tariffData.find((tr) => {
                  return (
                    equivalentOrderSkus.includes(
                      tr[tSkuIdx]
                        ? String(tr[tSkuIdx]).trim().toUpperCase()
                        : "",
                    ) &&
                    (tr[tCountryIdx]
                      ? String(tr[tCountryIdx]).trim().toUpperCase()
                      : "") === country.toUpperCase()
                  );
                });

                if (tariffRow) {
                  const baseSku = sku.toUpperCase();
                  const pInfo =
                    productInfoMap[baseSku] ||
                    getEquivalentSkus(baseSku)
                      .map((s) => productInfoMap[s])
                      .find(Boolean);

                  const rowLength = lengthIdx !== -1 ? parsePositiveNumber(tariffRow[lengthIdx]) : 0;
                  const rowWidth = widthIdx !== -1 ? parsePositiveNumber(tariffRow[widthIdx]) : 0;
                  const rowHeight = heightIdx !== -1 ? parsePositiveNumber(tariffRow[heightIdx]) : 0;
                  const rowWeight = weightIdx !== -1 ? parsePositiveNumber(tariffRow[weightIdx]) : 0;
                  const pL = rowLength || pInfo?.l || 0;
                  const pW = rowWidth || pInfo?.w || 0;
                  const pH = rowHeight || pInfo?.h || 0;
                  const pWT = rowWeight || pInfo?.wt || 0;

                  Object.keys(tWhs).forEach((whKey) => {
                    const whCols = tWhs[whKey];
                    whCols.allCols.forEach((col: any) => {
                      const valStr = tariffRow[col.idx];
                      const val = parseNumber(valStr, NaN);
                      if (!isNaN(val) && val > 0 && val < lowestCost) {
                        if (!isChannelShippable(col.name, country, pL, pW, pH, pWT))
                          return;

                        lowestCost = val;
                        bestPriceWH = whKey;
                        bestPriceChannel = col.name;
                      }
                    });
                  });
                }
              }

              if (lowestCost !== Infinity) {
                costCalcCount++;
                row[cheapestWhCol] = bestPriceWH;
                row[cheapestChCol] = bestPriceChannel;
                row[costColIdx] = lowestCost;

                const combinedStock = getCombinedInventory(sku);
                const hasStock = combinedStock.some(
                  (s) => s.avail > 0 && isSameWarehouse(s.wh, bestPriceWH),
                );
                row[stockStatusCol] = hasStock ? "现货满足" : "需调拨/无货";
                row[suggestedSkuCol] = hasStock
                  ? getSuggestedStockSku(sku, bestPriceWH) || "-"
                  : "-";
              } else {
                row[cheapestWhCol] = "无资费数据";
                row[cheapestChCol] = "-";
                row[costColIdx] = "-";
                row[stockStatusCol] = "-";
                row[suggestedSkuCol] = "-";
              }
            }
            wb.Sheets[sheetNames[2]] = XLSX.utils.aoa_to_sheet(data3);
            addLog(
              `完成 Sheet 3 [订单校验]: 改单建议 ${sheet3Updated} 条。`,
              "success",
            );
            if (costCalcCount > 0) {
              addLog(
                `完成 Sheet 3 [资费最优核算]: 成功挂载 ${costCalcCount} 条最低运费比价结果，已写入末尾新列。`,
                "success",
              );
            } else if (tariffData) {
              addLog(
                `警告: [资费最优核算] 尝试执行，但在报价表中未能匹配相应的 SKU+国家 数据。`,
                "error",
              );
            }
          }
        } else {
          addLog(
            "未检测到 Sheet 3 (订单明细表)，已跳过换仓改单逻辑。",
            "error",
          );
        }

        if (sheetNames.length < 2) {
          addLog(
            `核心警告: 当前上传的工作簿只有 ${sheetNames.length} 个工作表。系统需要 Sheet2 提供库龄用于推断，且需要 Sheet3 作为订单来源进行修改。处理的数据可能不完整。`,
            "error",
          );
        } else {
          addLog(`所有策略扫描完毕！已准备好合成新的 Excel 表格。`, "success");
        }

        setProcessedWb(wb);
      } catch (err: any) {
        addLog(`全局解析错误导致中止: ${err.message}`, "error");
      } finally {
        setLoading(false);
      }
    }, 300);
  };

  const handleDownload = () => {
    if (processedWb && fileData) {
      XLSX.writeFile(processedWb, `发货切分指导_${fileData.name}`);
    }
  };

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="bg-indigo-900 px-6 py-4 border-b border-indigo-800 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-500/20 p-2 rounded-lg border border-indigo-400/30">
            <Rocket className="w-5 h-5 text-indigo-300" />
          </div>
          <div>
            <h2 className="text-white font-bold tracking-wide">
              智能订单防错与库龄策略路由
            </h2>
            <p className="text-indigo-300 text-xs mt-0.5">
              自动解析 T2/粮仓 拦截指令，依据库龄推断最优海外仓切分方案
            </p>
          </div>
        </div>
      </div>
      <div className="p-6 sm:p-8 space-y-8">
        <div>
          <h3 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-2 uppercase tracking-wide">
            <span className="bg-indigo-100 text-indigo-700 w-6 h-6 flex items-center justify-center rounded-full text-xs">
              1
            </span>
            导入联动作业工作簿
          </h3>
          <div
            className={`border-2 border-dashed rounded-xl p-8 text-center transition-colors cursor-pointer ${
              fileData
                ? "border-emerald-300 bg-emerald-50/50 hover:bg-emerald-50"
                : "border-slate-300 hover:border-indigo-400 bg-slate-50"
            }`}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              type="file"
              accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
              className="hidden"
              ref={fileInputRef}
              onChange={handleUpload}
            />

            {fileData ? (
              <div className="text-emerald-700 flex flex-col items-center">
                <div className="bg-emerald-100 p-3 rounded-full mb-3">
                  <CheckSquare className="w-8 h-8 text-emerald-600" />
                </div>
                <p className="font-bold">源文件就绪</p>
                <p className="text-sm mt-1 text-emerald-600/70 max-w-md truncate font-mono">
                  {fileData.name}
                </p>
              </div>
            ) : (
              <div className="text-slate-500 flex flex-col items-center">
                <div className="bg-white p-3 rounded-full shadow-sm mb-3 border border-slate-100">
                  <Upload className="w-8 h-8 text-indigo-500" />
                </div>
                <p className="font-bold text-slate-700">
                  点击或拖入业务台账 (包含多Sheet工作簿)
                </p>
                <p className="text-sm mt-1 text-slate-400">
                  系统将自动读取 Sheet1(关注SKU) / Sheet2(库龄情况) /
                  Sheet3(欧盟订单明细) 联合运算
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <button
            onClick={processExcel}
            disabled={!fileData || loading}
            className="bg-slate-900 hover:bg-slate-800 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold py-4 rounded-xl flex items-center justify-center gap-2 transition-all shadow-md active:scale-[0.99]"
          >
            {loading ? (
              <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
            ) : (
              <GitMerge className="w-5 h-5" />
            )}
            执行融合路由算法
          </button>

          <button
            onClick={handleDownload}
            disabled={!processedWb}
            className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold py-4 rounded-xl flex items-center justify-center gap-2 transition-all shadow-md active:scale-[0.99]"
          >
            <FileDown className="w-5 h-5" />
            下载优化出库指导表
          </button>
        </div>

        {!tariffData && (
          <div className="bg-amber-50 border border-amber-200 text-amber-800 p-4 rounded-xl flex items-start gap-3">
            <AlertTriangle className="w-5 h-5 flex-shrink-0 mt-0.5 text-amber-500" />
            <div className="text-xs">
              <span className="font-bold block mb-1 text-sm text-amber-900">
                未检测到全局资费矩阵
              </span>
              您尚未在「资费解析」模块读取报价表。系统依然可以执行库龄拦截和欧盟防错换仓，但无法执行
              <strong>「最优运费与渠道」全局比价核算功能</strong>
              。如需附带该功能，请先在资费模块导入报价库。
            </div>
          </div>
        )}

        {logs.length > 0 && (
          <div className="bg-slate-900 rounded-xl p-5 shadow-inner">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-700 pb-3">
              <Database className="w-4 h-4 text-indigo-400" />
              <span className="text-xs font-mono font-bold text-slate-300 uppercase tracking-widest">
                Execution Trace
              </span>
            </div>
            <div className="space-y-2.5 max-h-64 overflow-y-auto font-mono text-[11.5px] leading-relaxed pr-2">
              {logs.map((log) => (
                <div key={log.id} className="flex gap-3">
                  <span className="text-slate-500 shrink-0">
                    [{new Date().toLocaleTimeString()}]
                  </span>
                  <span
                    className={`
                    ${
                      log.type === "error"
                        ? "text-rose-400"
                        : log.type === "success"
                          ? "text-emerald-400"
                          : "text-indigo-200"
                    }`}
                  >
                    {log.msg}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState<"query" | "calc" | "orders">(
    "query",
  );

  // === 模块 1: 运费查询状态 ===
  const [fileType, setFileType] = useState("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [currentSheet, setCurrentSheet] = useState("");

  const [csvData, setCsvData] = useState<any[][] | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState("");

  const [meta, setMeta] = useState({
    skuIndex: -1,
    countryIndex: -1,
    lengthIdx: -1,
    widthIdx: -1,
    heightIdx: -1,
    weightIdx: -1,
    categoryIdx: -1,
    warehouses: {} as Record<string, any>,
    skuList: [] as string[],
    countryList: [] as string[],
    warehouseList: [] as string[],
  });

  const [inputSku, setInputSku] = useState("");
  const [inputCountry, setInputCountry] = useState("");
  const [result, setResult] = useState<any>(null);
  const [showAllCosts, setShowAllCosts] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 2. 优先发货与库存数据
  const [priorityMap, setPriorityMap] = useState<Record<string, any>>({});
  const [inventoryMap, setInventoryMap] = useState<
    Record<string, Record<string, number>>
  >({});
  const [priorityFile, setPriorityFile] = useState<string | null>(null);
  const [priorityLoading, setPriorityLoading] = useState(false);
  const [isEUAccount, setIsEUAccount] = useState(false);

  // 3. 产品资料数据
  const [productInfoMap, setProductInfoMap] = useState<
    Record<
      string,
      { l: number; w: number; h: number; wt: number; qty?: number }
    >
  >({});
  const [productFile, setProductFile] = useState<string | null>(null);
  const [productLoading, setProductLoading] = useState(false);
  const productInputRef = useRef<HTMLInputElement>(null);

  // 展开明细的状态控制
  const [showBestOptionDetails, setShowBestOptionDetails] = useState(false);
  const [expandedAltWH, setExpandedAltWH] = useState<string | null>(null);

  const priorityInputRef = useRef<HTMLInputElement>(null);

  // === 模块 2: 尺寸计算状态 ===
  const [inputs, setInputs] = useState({
    length: 26,
    width: 18.5,
    height: 15.5,
    weight: 0.5,
    qty: 5,
  });
  const [stackAxis, setStackAxis] = useState("min");

  useEffect(() => {
    async function loadStoredData() {
      try {
        const storedCsvData = await localforage.getItem("csvData");
        if (storedCsvData) setCsvData(storedCsvData as any);

        const storedMeta = await localforage.getItem("meta");
        if (storedMeta) setMeta(storedMeta as any);

        const storedFileName = await localforage.getItem("fileName");
        if (storedFileName) setFileName(storedFileName as any);

        const storedPriorityMap = await localforage.getItem("priorityMap");
        if (storedPriorityMap) setPriorityMap(storedPriorityMap as any);

        const storedInventoryMap = await localforage.getItem("inventoryMap");
        if (storedInventoryMap) setInventoryMap(storedInventoryMap as any);

        const storedPriorityFile = await localforage.getItem("priorityFile");
        if (storedPriorityFile) setPriorityFile(storedPriorityFile as any);
        
        const storedIsEu = await localforage.getItem("isEUAccount");
        if (storedIsEu !== null) setIsEUAccount(storedIsEu as boolean);
        
        const storedProductInfoMap = await localforage.getItem("productInfoMap");
        if (storedProductInfoMap) setProductInfoMap(storedProductInfoMap as any);
        
        const storedProductFile = await localforage.getItem("productFile");
        if (storedProductFile) setProductFile(storedProductFile as any);

      } catch (e) {
        console.error("加载本地数据失败", e);
      }
    }
    loadStoredData();
  }, []);

  // --- 数据表严谨解析 ---
  const processTableData = useCallback((parsed: any[][]) => {
    if (!parsed || parsed.length < 3)
      throw new Error("工作表内容行数过少或为空，请确认文件完整性。");

    const row0 = parsed[0];

    // 智能探测真实的底层表头行
    let subHeaderIdx = 1;
    for (let i = 2; i <= 4 && i < parsed.length; i++) {
      const rowArr = parsed[i];
      if (!rowArr) continue;
      const hasHeaderSign = rowArr.some((c) => {
        const s = String(c).trim().toUpperCase();
        return s === "SKU" || s.includes("最佳渠道") || s.includes("最佳运费");
      });
      if (hasHeaderSign) {
        subHeaderIdx = i;
      }
    }

    const mainHeaderRow = parsed[1] || [];
    const subHeaderRow = parsed[subHeaderIdx] || [];
    const dataRows = parsed
      .slice(subHeaderIdx + 1)
      .filter(
        (row) => row && row.length > 0 && row.some((cell) => cell !== ""),
      );

    const getCellStr = (row: any[], idx: number) =>
      row[idx] ? String(row[idx]).trim() : "";

    const matchHeader = (keywords: string[]) => {
      let idx = subHeaderRow.findIndex((col) =>
        keywords.some(
          (kw) => String(col).trim().toUpperCase() === kw.toUpperCase(),
        ),
      );
      if (idx !== -1) return idx;

      if (subHeaderIdx !== 1) {
        idx = mainHeaderRow.findIndex((col) =>
          keywords.some(
            (kw) => String(col).trim().toUpperCase() === kw.toUpperCase(),
          ),
        );
        if (idx !== -1) return idx;
      }

      idx = subHeaderRow.findIndex((col) => {
        const val = String(col).trim().toUpperCase();
        if (
          val.includes("*") ||
          val.includes("&") ||
          val.includes("数量") ||
          val.includes("组合") ||
          val.includes("+")
        )
          return false;
        return keywords.some((kw) => val.includes(kw.toUpperCase()));
      });
      if (idx !== -1) return idx;

      if (subHeaderIdx !== 1) {
        idx = mainHeaderRow.findIndex((col) => {
          const val = String(col).trim().toUpperCase();
          if (
            val.includes("*") ||
            val.includes("&") ||
            val.includes("数量") ||
            val.includes("组合") ||
            val.includes("+")
          )
            return false;
          return keywords.some((kw) => val.includes(kw.toUpperCase()));
        });
        if (idx !== -1) return idx;
      }

      return -1;
    };

    const skuIndex = matchHeader(["SKU", "SKU编码", "商品编码"]);
    const countryIndex = matchHeader([
      "国家",
      "目的国",
      "COUNTRY",
      "DESTINATION",
    ]);
    const lengthIdx = matchHeader(["长", "LENGTH"]);
    const widthIdx = matchHeader(["宽", "WIDTH"]);
    const heightIdx = matchHeader(["高", "HEIGHT"]);
    const weightIdx = matchHeader(["实重", "重量", "WEIGHT"]);
    const categoryIdx = matchHeader([
      "品类",
      "产品类别",
      "CATEGORY",
      "产品品类",
    ]);

    if (skuIndex === -1 || countryIndex === -1) {
      throw new Error(
        "未能精确匹配到'SKU'或'国家'标识列，请检查表头是否符合规范规范。",
      );
    }

    const skipKeywords = [
      "SKU",
      "编码",
      "国家",
      "COUNTRY",
      "DESTINATION",
      "长",
      "LENGTH",
      "宽",
      "WIDTH",
      "高",
      "HEIGHT",
      "重",
      "WEIGHT",
      "体积",
      "利润",
      "成本",
      "最佳",
      "次佳",
      "备选",
      "辅助列",
    ];

    const warehouses: Record<string, any> = {};
    let currentWarehouse = "";

    for (let i = 0; i < row0.length; i++) {
      const cell0 = getCellStr(row0, i);
      const cell1 = getCellStr(mainHeaderRow, i);
      const cellLast = getCellStr(subHeaderRow, i);

      if (
        cell0 &&
        !["辅助列"].includes(cell0) &&
        !cell0.includes("SKU*数量") &&
        !skipKeywords.some((kw) => cell0.toUpperCase().includes(kw))
      ) {
        currentWarehouse = cell0;
        if (!warehouses[currentWarehouse])
          warehouses[currentWarehouse] = {
            channelCol: -1,
            costCol: -1,
            subChannelCol: -1,
            subCostCol: -1,
            allCols: [],
          };
      }

      if (currentWarehouse && warehouses[currentWarehouse]) {
        const combinedHeader = cell1 + " " + cellLast;
        if (combinedHeader.includes("最佳渠道"))
          warehouses[currentWarehouse].channelCol = i;
        else if (combinedHeader.includes("最佳运费"))
          warehouses[currentWarehouse].costCol = i;
        else if (
          combinedHeader.includes("次佳渠道") ||
          combinedHeader.includes("备选渠道")
        )
          warehouses[currentWarehouse].subChannelCol = i;
        else if (
          combinedHeader.includes("次佳运费") ||
          combinedHeader.includes("备选运费")
        )
          warehouses[currentWarehouse].subCostCol = i;

        const specificChannelName = cellLast || cell1;

        if (
          specificChannelName &&
          !skipKeywords.some((kw) =>
            specificChannelName.toUpperCase().includes(kw),
          )
        ) {
          warehouses[currentWarehouse].allCols.push({
            name: specificChannelName,
            idx: i,
          });
        }
      }
    }

    const validWarehouses = Object.keys(warehouses).filter(
      (w) =>
        (warehouses[w].channelCol !== -1 && warehouses[w].costCol !== -1) ||
        warehouses[w].allCols.length > 0,
    );

    if (validWarehouses.length === 0)
      throw new Error(
        "未检测到有效的仓库费用列(缺少'最佳运费'或渠道价格数据)。",
      );

    const skuSet = new Set<string>();
    const countrySet = new Set<string>();

    dataRows.forEach((row) => {
      const sku = getCellStr(row, skuIndex);
      const country = getCellStr(row, countryIndex);
      if (sku) skuSet.add(sku);
      if (country) countrySet.add(country);
    });

    const newMeta = {
      skuIndex,
      countryIndex,
      lengthIdx,
      widthIdx,
      heightIdx,
      weightIdx,
      categoryIdx,
      warehouses,
      skuList: Array.from(skuSet).sort(),
      countryList: Array.from(countrySet).sort(),
      warehouseList: validWarehouses,
    };
    
    setMeta(newMeta);
    setCsvData(dataRows);
    localforage.setItem("meta", newMeta);
    localforage.setItem("csvData", dataRows);
  }, []);

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value;
    setCurrentSheet(sheetName);
    setError(null);
    setResult(null);
    try {
      if (workbook && sheetName) {
        const worksheet = workbook.Sheets[sheetName];
        const parsed = XLSX.utils.sheet_to_json<any[]>(worksheet, {
          header: 1,
          defval: "",
        });
        processTableData(parsed);
      }
    } catch (err: any) {
      setError(`工作表解析异常: ${err.message}`);
      setCsvData(null);
    }
  };

  const processPriorityData = (
    sheetsData: { name: string; data: any[][] }[],
  ) => {
    const newPriorityMap = { ...priorityMap };
    const newInvMap = { ...inventoryMap };

    const parseSheet = (
      parsed: any[][],
      sheetName: string,
      expectedType: "t2" | "age",
    ) => {
      if (!parsed || parsed.length < 2) return;
      const headers = parsed[0] || [];
      const getCell = (row: any[], idx: number) =>
        row[idx] ? String(row[idx]).trim() : "";

      const skuIdx = headers.findIndex(
        (h: any) => typeof h === "string" && h.trim().toUpperCase() === "SKU",
      );
      const cleanSkuIdx = headers.findIndex(
        (h: any) => typeof h === "string" && h.includes("剔除前序"),
      );
      const whIdx = headers.findIndex(
        (h: any) => typeof h === "string" && h.trim() === "仓库",
      );
      // For age sheet, explicitly look for these
      const stockIdx = headers.findIndex(
        (h: any) =>
          typeof h === "string" &&
          (h.trim() === "可用库存" ||
            h.trim() === "库存" ||
            h.includes("可用")),
      );

      if (skuIdx === -1 || whIdx === -1) return;

      const priorityType = expectedType;

      for (let i = 1; i < parsed.length; i++) {
        const row = parsed[i];
        const sku = skuIdx !== -1 ? getCell(row, skuIdx).toUpperCase() : "";
        const cleanSku =
          cleanSkuIdx !== -1 ? getCell(row, cleanSkuIdx).toUpperCase() : "";
        const wh = getCell(row, whIdx);

        const stockStr = stockIdx !== -1 ? getCell(row, stockIdx) : "";
        const stockNum = parseNumber(stockStr);

        if (wh && wh !== "/") {
          if (sku) {
            if (!newPriorityMap[sku]) newPriorityMap[sku] = {};
            newPriorityMap[sku][priorityType] = wh;
          }
          if (cleanSku) {
            if (!newPriorityMap[cleanSku]) newPriorityMap[cleanSku] = {};
            newPriorityMap[cleanSku][priorityType] = wh;
          }

          if (priorityType === "age") {
            if (sku) {
              if (!newInvMap[sku]) newInvMap[sku] = {};
              newInvMap[sku][wh] = Math.max(newInvMap[sku][wh] || 0, stockNum);
            }
            if (cleanSku) {
              if (!newInvMap[cleanSku]) newInvMap[cleanSku] = {};
              newInvMap[cleanSku][wh] = Math.max(
                newInvMap[cleanSku][wh] || 0,
                stockNum,
              );
            }
          }
        }
      }
    };

    if (sheetsData.length > 0) {
      parseSheet(sheetsData[0].data, sheetsData[0].name, "t2"); // Sheet 1 is always T2
    }
    if (sheetsData.length > 1) {
      parseSheet(sheetsData[1].data, sheetsData[1].name, "age"); // Sheet 2 is always Age/Stock
    }

    setPriorityMap(newPriorityMap);
    setInventoryMap(newInvMap);
    localforage.setItem("priorityMap", newPriorityMap);
    localforage.setItem("inventoryMap", newInvMap);
  };

  const processProductData = useCallback((parsed: any[][]) => {
    if (!parsed || parsed.length < 2) return;

    // 探测真实的表头行
    let headerIdx = 0;
    for (let i = 0; i < Math.min(5, parsed.length); i++) {
      if (
        parsed[i] &&
        parsed[i].some((col) => {
          const s = String(col).trim().toUpperCase();
          return s.includes("SKU") || s.includes("编码");
        })
      ) {
        headerIdx = i;
        break;
      }
    }

    const headerRow = parsed[headerIdx] || [];
    const getIdx = (kws: string[]) => {
      for (const k of kws) {
        const exactIdx = headerRow.findIndex(
          (c) => String(c).trim().toUpperCase() === k.toUpperCase(),
        );
        if (exactIdx !== -1) return exactIdx;
      }
      for (const k of kws) {
        const idx = headerRow.findIndex((c) =>
          String(c).trim().toUpperCase().includes(k.toUpperCase()),
        );
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const skuCols: number[] = [];
    for (let c = 0; c < headerRow.length; c++) {
      const colStr = String(headerRow[c]).trim().toUpperCase();
      if (
        ["通用SKU", "亚马逊SKU", "EBAYSKU", "SKU", "商品编码", "编码"].includes(
          colStr,
        ) ||
        colStr.includes("SKU")
      ) {
        skuCols.push(c);
      }
    }

    const lIdx = getIdx(["长cm", "长", "LENGTH"]);
    const wIdx = getIdx(["宽cm", "宽", "WIDTH"]);
    const hIdx = getIdx(["高cm", "高", "HEIGHT"]);
    const wtIdx = getIdx([
      "总发毛重 kg",
      "总发毛重",
      "实重",
      "重量",
      "WEIGHT",
      "毛重",
      "净重",
    ]);
    const qtyIdx = getIdx([
      "一箱件数",
      "一箱装数量",
      "装箱数量",
      "数量",
      "QTY",
      "QUANTITY",
    ]);

    if (skuCols.length === 0) {
      throw new Error("产品箱规表中缺少商品编码(SKU)列。请检查表头。");
    }

    const newMap: Record<
      string,
      { l: number; w: number; h: number; wt: number; qty?: number }
    > = {};
    let validCount = 0;
    for (let i = headerIdx + 1; i < parsed.length; i++) {
      const row = parsed[i];
      if (!row) continue;

      let hasSku = false;
      const skusToMap: string[] = [];
      for (const c of skuCols) {
        if (row[c]) {
          const skuStr = String(row[c]).trim().toUpperCase();
          if (skuStr) {
            getEquivalentSkus(skuStr).forEach((sku) => skusToMap.push(sku));
            hasSku = true;
          }
        }
      }

      if (!hasSku) continue;

      const l = parsePositiveNumber(row[lIdx]);
      const w = parsePositiveNumber(row[wIdx]);
      const h = parsePositiveNumber(row[hIdx]);
      const wt = parsePositiveNumber(row[wtIdx]);

      let qty: number | undefined = undefined;
      if (qtyIdx !== -1 && row[qtyIdx] !== undefined) {
        const qtyParsed = Math.trunc(parseNumber(row[qtyIdx], NaN));
        if (!isNaN(qtyParsed) && qtyParsed > 0) {
          qty = qtyParsed;
        }
      }

      if (l > 0 || wt > 0 || w > 0 || h > 0 || qty !== undefined) {
        Array.from(new Set(skusToMap)).forEach((sku) => {
          newMap[sku] = { l, w, h, wt, qty };
        });
        validCount++;
      }
    }

    if (validCount === 0) {
      throw new Error("未能读取到有效的SKU或数量数据，可能是列名不匹配或数据为空。请使用下载按钮获取模板查看标准格式！");
    }

    setProductInfoMap(newMap);
    localforage.setItem("productInfoMap", newMap);
  }, []);

  const downloadProductTemplate = (e: React.MouseEvent) => {
    e.stopPropagation();
    const content =
      "通用sku,亚马逊SKU,eBaySKU,一箱件数,长cm,宽cm,高cm,总发毛重 kg\nC-1-HG4028A,AM-C-1-HG4028A,EB-C-1-HG4028A,1,10,10,10,1.5\nC-1-HG4029A,AM-C-1-HG4029A,EB-C-1-HG4029A,2,20,15,10,2";
    const blob = new Blob([new Uint8Array([0xef, 0xbb, 0xbf]), content], {
      type: "text/csv;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "产品箱规导入模板.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>,
    type: "freight" | "priority" | "product" = "freight",
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setError(null);
    const isExcel = file.name.match(/\.(xlsx|xls)$/i);

    if (type === "freight") {
      setFileName(file.name);
      localforage.setItem("fileName", file.name);
      setLoading(true);
    } else if (type === "priority") {
      setPriorityFile(file.name);
      localforage.setItem("priorityFile", file.name);
      setPriorityLoading(true);
    } else if (type === "product") {
      setProductFile(file.name);
      localforage.setItem("productFile", file.name);
      setProductLoading(true);
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = event.target?.result;
        if (!data) throw new Error("文件读取失败");

        const wb = XLSX.read(data, { type: "array" });

        if (type === "freight") {
          const targetSheetName =
            wb.SheetNames.find((s) => s.includes("尾程")) || wb.SheetNames[0];
          const parsed = XLSX.utils.sheet_to_json<any[]>(
            wb.Sheets[targetSheetName],
            { header: 1, defval: "" },
          );
          processTableData(parsed);
        } else if (type === "priority") {
          const sheetsData = wb.SheetNames.map((name) => ({
            name,
            data: XLSX.utils.sheet_to_json<any[]>(wb.Sheets[name], {
              header: 1,
              defval: "",
            }),
          }));
          processPriorityData(sheetsData);
        } else if (type === "product") {
          const targetSheetName = wb.SheetNames[0]; // Assuming first sheet for product data
          const parsed = XLSX.utils.sheet_to_json<any[]>(
            wb.Sheets[targetSheetName],
            { header: 1, defval: "" },
          );
          processProductData(parsed);
        }
      } catch (err: any) {
        setError(
          `[${type === "freight" ? "尾程" : type === "priority" ? "优先" : "产品箱规"}] 文件解析错误: ` +
            err.message,
        );
        window.scrollTo({ top: 0, behavior: "smooth" });
      } finally {
        if (type === "freight") setLoading(false);
        else if (type === "priority") setPriorityLoading(false);
        else if (type === "product") setProductLoading(false);
      }
    };
    reader.onerror = () => {
      setError("底层文件读取被阻断或文件损坏");
      if (type === "freight") setLoading(false);
      else if (type === "priority") setPriorityLoading(false);
      else if (type === "product") setProductLoading(false);
    };

    reader.readAsArrayBuffer(file);

    e.target.value = "";
  };

  const handleClearCache = async () => {
    try {
      await localforage.clear();
      setCsvData(null);
      setMeta({
        skuIndex: -1, countryIndex: -1, lengthIdx: -1, widthIdx: -1, heightIdx: -1, weightIdx: -1, categoryIdx: -1,
        warehouses: {}, skuList: [], countryList: [], warehouseList: []
      });
      setFileName("");
      setPriorityMap({});
      setInventoryMap({});
      setPriorityFile(null);
      setProductInfoMap({});
      setProductFile(null);
      setResult(null);
      setError(null);
    } catch (e) {
      console.error("清空缓存失败", e);
    }
  };

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    if (!csvData) {
      setError("请先上传尾程报价表。");
      return;
    }

    setShowBestOptionDetails(false);
    setExpandedAltWH(null);
    setError(null);
    setResult(null);

    const cleanSku = inputSku.trim().toUpperCase();
    const cleanCountry = inputCountry.trim();
    if (!cleanSku || !cleanCountry) {
      setError("SKU和国家为必填项。");
      return;
    }
    const equivalentInputSkus = getEquivalentSkus(cleanSku);

    const {
      skuIndex,
      countryIndex,
      warehouses,
      lengthIdx,
      widthIdx,
      heightIdx,
      weightIdx,
    } = meta;

    const getCellStr = (row: any[], idx: number) =>
      row[idx] ? String(row[idx]).trim() : "";

    const matchRow = csvData.find((row) => {
      return (
        equivalentInputSkus.includes(getCellStr(row, skuIndex).toUpperCase()) &&
        getCellStr(row, countryIndex) === cleanCountry
      );
    });

    if (!matchRow) {
      setError(`未找到 SKU: ${cleanSku} 且发往 ${cleanCountry} 的记录。`);
      return;
    }

    const hasInventoryData = Object.keys(inventoryMap).length > 0;

    const extractCostValue = (str: string) => {
      const clean = String(str ?? "").trim();
      if (!clean) return Infinity;
      if (clean.startsWith("#") || clean.includes("N/A") || clean === "-")
        return Infinity;
      const val = parseNumber(clean, Infinity);
      return Number.isFinite(val) ? val : Infinity;
    };

    const getStock = (skuKey: string, targetWh: string) => {
      const eqSkus = getEquivalentSkus(skuKey);
      let ebStock = 0;
      let amStock = 0;
      let t1Stock = 0;
      let t2Stock = 0;
      let rawStock = 0;
      let found = false;
      const stockBySku: Record<string, number> = {};
      eqSkus.forEach((s) => {
        if (inventoryMap[s]) {
          let sQty = 0;
          for (const whKey in inventoryMap[s]) {
            if (targetWh.includes(whKey) || whKey.includes(targetWh)) {
              sQty += inventoryMap[s][whKey];
              found = true;
            }
          }
          if (s.startsWith("EB-")) ebStock += sQty;
          else if (s.startsWith("AM-")) amStock += sQty;
          else if (s.startsWith("T1-")) t1Stock += sQty;
          else if (s.startsWith("T2-")) t2Stock += sQty;
          else rawStock += sQty;
          if (sQty > 0) stockBySku[s] = (stockBySku[s] || 0) + sQty;
        }
      });
      const preferredSkuOrder = [
        skuKey.trim().toUpperCase(),
        getSkuWithPrefix(skuKey, "AM"),
        getSkuWithPrefix(skuKey, "EB"),
        getSkuWithPrefix(skuKey, "T1"),
        getSkuWithPrefix(skuKey, "T2"),
        getSkuBase(skuKey),
        ...eqSkus,
      ];
      const suggestedSku = Array.from(new Set(preferredSkuOrder)).find(
        (s) => (stockBySku[s] || 0) > 0,
      );
      return found
        ? {
            total: ebStock + amStock + t1Stock + t2Stock + rawStock,
            eb: ebStock,
            am: amStock,
            t1: t1Stock,
            t2: t2Stock,
            raw: rawStock,
            suggestedSku,
          }
        : undefined;
    };

    const pInfo =
      productInfoMap[cleanSku] ||
      getEquivalentSkus(cleanSku)
        .map((s) => productInfoMap[s])
        .find(Boolean);
    const rowLength = lengthIdx !== -1 ? parsePositiveNumber(getCellStr(matchRow, lengthIdx)) : 0;
    const rowWidth = widthIdx !== -1 ? parsePositiveNumber(getCellStr(matchRow, widthIdx)) : 0;
    const rowHeight = heightIdx !== -1 ? parsePositiveNumber(getCellStr(matchRow, heightIdx)) : 0;
    const rowWeight = weightIdx !== -1 ? parsePositiveNumber(getCellStr(matchRow, weightIdx)) : 0;
    const pL = rowLength || pInfo?.l || 0;
    const pW = rowWidth || pInfo?.w || 0;
    const pH = rowHeight || pInfo?.h || 0;
    const pWT = rowWeight || pInfo?.wt || 0;

    const options: any[] = [];
    Object.keys(warehouses).forEach((whName) => {
      const { channelCol, costCol, allCols } = warehouses[whName];
      // Note: we can't blindly trust the pre-calculated bestChannel if it's currently restricted
      // We will re-determine the best channel from details

      const details: any[] = [];
      allCols.forEach((ch: any) => {
        const cStr = getCellStr(matchRow, ch.idx);
        if (
          cStr &&
          !cStr.startsWith("#") &&
          !cStr.includes("N/A") &&
          cStr !== ""
        ) {
          const cNum = extractCostValue(cStr);
          if (cNum !== Infinity) {
            if (isChannelShippable(ch.name, cleanCountry, pL, pW, pH, pWT)) {
              details.push({ channel: ch.name, costStr: cStr, costNum: cNum });
            }
          }
        }
      });
      details.sort((a, b) => a.costNum - b.costNum);

      let effectiveBestChannel = "";
      let effectiveBestCostStr = "";
      let effectiveBestCostNum = Infinity;

      if (details.length > 0) {
        effectiveBestChannel = details[0].channel;
        effectiveBestCostStr = details[0].costStr;
        effectiveBestCostNum = details[0].costNum;
      }

      if (effectiveBestChannel && effectiveBestCostNum !== Infinity) {
        let currentStock: any = "未知";
        if (hasInventoryData) {
          const stockVal = getStock(cleanSku, whName);
          currentStock =
            stockVal !== undefined
              ? stockVal
              : { total: 0, eb: 0, am: 0, t1: 0, t2: 0, raw: 0 };
        }

        options.push({
          warehouse: whName,
          channel: effectiveBestChannel,
          costStr: effectiveBestCostStr,
          costNum: effectiveBestCostNum,
          stock: currentStock,
          details,
        });
      }
    });

    if (options.length === 0) {
      setError(
        `找到关联记录，但所有的发运仓库/渠道为空，或已被渠道尺寸/重量限制规则自动拦截。`,
      );
      return;
    }

    const thirdPartyWHs = ["谷仓", "易达", "捎客", "大麦"];

    let validOptions = options;
    if (hasInventoryData) {
      validOptions = options.filter(
        (o) => typeof o.stock === "object" && o.stock.total > 0,
      );
    }

    let bestOption = null;
    let recommendReason = "";
    let recommendType = "cost";

    if (validOptions.length === 0 && hasInventoryData) {
      bestOption = {
        warehouse: "/",
        channel: "拦截：所有仓均无可用库存",
        costStr: "N/A",
        costNum: Infinity,
        details: [],
        stock: { total: 0, eb: 0, am: 0, t1: 0, t2: 0, raw: 0 },
      };
      recommendReason = "库存不足拦截 (全仓缺货)";
      recommendType = "no-stock";
    } else if (isEUAccount) {
      const thirdPartyOptions = validOptions.filter((o) =>
        thirdPartyWHs.some((tw) => o.warehouse.includes(tw)),
      );

      if (thirdPartyOptions.length > 0) {
        let pData: any = {};
        getEquivalentSkus(cleanSku).forEach((s) => {
          if (priorityMap[s]) pData = { ...pData, ...priorityMap[s] };
        });
        const ageWH = pData.age;

        let matchedTP = null;
        if (ageWH && thirdPartyWHs.some((tw) => ageWH.includes(tw))) {
          matchedTP = thirdPartyOptions.find((o) =>
            o.warehouse.includes(ageWH),
          );
        }

        if (matchedTP) {
          bestOption = matchedTP;
          recommendReason = `欧盟账号 + 有库存 + 库龄优先 (${ageWH})`;
          recommendType = "eu-priority";
        } else {
          thirdPartyOptions.sort((a, b) => a.costNum - b.costNum);
          bestOption = thirdPartyOptions[0];
          recommendReason = `欧盟限制：分配有库存的第三方仓 (运费最优)`;
          recommendType = "eu-cost";
        }
      } else {
        bestOption = {
          warehouse: "/",
          channel: "拦截：无可用或有货的第三方仓",
          costStr: "N/A",
          costNum: Infinity,
          details: [],
          stock: { total: 0, eb: 0, am: 0, t1: 0, t2: 0, raw: 0 },
        };
        recommendReason = "欧盟强限制：第三方仓配置不足或缺货";
        recommendType = "eu-blocked";
      }
    } else {
      validOptions.sort((a, b) => a.costNum - b.costNum);
      let pData: any = {};
      getEquivalentSkus(cleanSku).forEach((s) => {
        if (priorityMap[s]) pData = { ...pData, ...priorityMap[s] };
      });
      const targetWH = pData.age || pData.t2;

      if (targetWH) {
        const matchedOption = validOptions.find(
          (o) =>
            o.warehouse.includes(targetWH) || targetWH.includes(o.warehouse),
        );
        if (matchedOption) {
          bestOption = matchedOption;
          recommendReason = pData.age
            ? `库龄预警 + 库存充足 (${targetWH})`
            : `T2/通用优先 + 库存充足 (${targetWH})`;
          recommendType = "priority";
        }
      }

      if (!bestOption) {
        bestOption = validOptions[0];
        recommendReason = "寻源有库存的最优运费仓";
        recommendType = "cost";
      }
    }

    const itemL =
      lengthIdx !== -1 ? parseNumber(getCellStr(matchRow, lengthIdx), NaN) : NaN;
    const itemW =
      widthIdx !== -1 ? parseNumber(getCellStr(matchRow, widthIdx), NaN) : NaN;
    const itemH =
      heightIdx !== -1 ? parseNumber(getCellStr(matchRow, heightIdx), NaN) : NaN;
    const itemWt =
      weightIdx !== -1 ? parseNumber(getCellStr(matchRow, weightIdx), NaN) : NaN;

    const hasDims =
      (!isNaN(itemL) && !isNaN(itemW) && !isNaN(itemH)) ||
      (pL > 0 && pW > 0 && pH > 0);
    const hasWt = !isNaN(itemWt) || pWT > 0;

    const finalL = !isNaN(itemL) ? itemL : pL;
    const finalW = !isNaN(itemW) ? itemW : pW;
    const finalH = !isNaN(itemH) ? itemH : pH;
    const finalWt = !isNaN(itemWt) ? itemWt : pWT;
    const finalQty = pInfo?.qty;

    if (hasDims && hasWt) {
      setInputs({
        length: finalL,
        width: finalW,
        height: finalH,
        weight: finalWt,
        qty: finalQty || 1,
      });
    }

    setResult({
      sku: cleanSku,
      country: cleanCountry,
      bestOption,
      recommendReason,
      recommendType,
      hasInventoryData,
      altOptions: options.filter(
        (o) => bestOption && o.warehouse !== bestOption.warehouse,
      ),
      dims: hasDims ? `${finalL} × ${finalW} × ${finalH} cm` : "未录入",
      weightVal: hasWt ? `${finalWt} kg` : "未录入",
      hasFullDims: hasDims && hasWt,
      hasQty: finalQty !== undefined,
      qtyVal: finalQty,
      rawDims: {
        l: finalL,
        w: finalW,
        h: finalH,
        wt: finalWt,
        qty: finalQty || 1,
      },
    });
  };

  const handleLinkToCalc = () => {
    if (result?.hasFullDims) {
      setInputs({
        length: result.rawDims.l,
        width: result.rawDims.w,
        height: result.rawDims.h,
        weight: result.rawDims.wt,
        qty: result.rawDims.qty || 1,
      });
      setActiveTab("calc");
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const handleCalcInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    const val = value === "" ? "" : Math.max(0, parseNumber(value, NaN));
    setInputs((prev) => ({ ...prev, [name]: isNaN(val as number) ? "" : val }));
  };

  const calcData = useMemo(() => {
    const { length, width, height, weight, qty } = inputs;

    const safeL = Number(length) || 0;
    const safeW = Number(width) || 0;
    const safeH = Number(height) || 0;
    const safeWt = Number(weight) || 0;
    const safeQty = Number(qty) || 1;

    const baseDims = [safeL, safeW, safeH].sort((a, b) => b - a);
    const baseMax = baseDims[0];
    const baseMid = baseDims[1];
    const baseMin = baseDims[2];

    let stackedMax = baseMax;
    let stackedMid = baseMid;
    let stackedMin = baseMin;
    if (stackAxis === "min") stackedMin = baseMin * safeQty;
    else if (stackAxis === "mid") stackedMid = baseMid * safeQty;
    else if (stackAxis === "max") stackedMax = baseMax * safeQty;

    const totalWeight = exactRound(safeWt * safeQty, 3);
    const finalDims = [stackedMax, stackedMid, stackedMin].sort(
      (a, b) => b - a,
    );
    const maxSide = exactRound(finalDims[0], 2);
    const midSide = exactRound(finalDims[1], 2);
    const minSide = exactRound(finalDims[2], 2);

    const girth = exactRound(maxSide + 2 * (midSide + minSide), 2);
    const volume = exactRound((maxSide * midSide * minSide) / 1000000, 4); // m3

    const checkDhl = () => {
      const isValid = maxSide <= 200 && girth <= 360 && totalWeight <= 31.5;
      const volWeight = exactRound((maxSide * midSide * minSide) / 5000, 2);
      const chargeWeight = Math.max(totalWeight, volWeight);

      const surcharges = [];
      if (maxSide > 120 || midSide > 60 || minSide > 60) {
        surcharges.push({ name: "境内超规附加费", fee: FREIGHT_FEES.DHL.domesticOversize });
      }
      if (maxSide < 15 || midSide < 11 || minSide < 1) {
        surcharges.push({ name: "低于最小尺寸附加费", fee: FREIGHT_FEES.DHL.domesticOversize });
      }

      return {
        isValid,
        volWeight,
        chargeWeight,
        surcharges,
        rules: {
          maxSide: maxSide <= 200,
          midSide: midSide <= 60,
          minSide: minSide <= 60,
          minLimit: maxSide >= 15 && midSide >= 11 && minSide >= 1,
          weight: totalWeight <= 31.5,
          girth: girth <= 360,
          volume: true,
        },
      };
    };

    const checkGls = () => {
      const isValid = true;
      const volWeight = exactRound((maxSide * midSide * minSide) / 6000, 2);
      const chargeWeight = Math.max(totalWeight, volWeight);

      const surcharges = [];
      if (volume > 0.15) {
        surcharges.push({ name: "超体积附加费", fee: FREIGHT_FEES.GLS.volume });
      }
      if (maxSide > 150) {
        surcharges.push({ name: "超长附加费", fee: FREIGHT_FEES.GLS.overLength });
      }

      let isNb = false;
      if (maxSide < 15 || midSide < 11 || minSide < 3) isNb = true;
      if (maxSide > 120 && maxSide <= 150 && midSide <= 80 && minSide <= 60)
        isNb = true;

      if (isNb) {
        surcharges.push({ name: "超尺寸(NB)附加费", fee: FREIGHT_FEES.GLS.nonConveyable });
      }

      if (
        maxSide > 200 ||
        midSide > 80 ||
        minSide > 60 ||
        girth >= 300 ||
        totalWeight > 40
      ) {
        surcharges.push({ name: "极限超规附加费", fee: FREIGHT_FEES.GLS.extreme });
      }

      return {
        isValid,
        volWeight,
        chargeWeight,
        surcharges,
        rules: {
          maxSide: maxSide <= 150,
          midSide: midSide <= 80,
          minSide: minSide <= 60,
          minLimit: maxSide >= 15 && midSide >= 11 && minSide >= 3,
          weight: totalWeight <= 40,
          girth: girth <= 300,
          volume: volume <= 0.15,
        },
      };
    };

    const checkDpd = () => {
      const isValid = true;
      const volWeight = exactRound((maxSide * midSide * minSide) / 5000, 2);
      const chargeWeight = Math.max(totalWeight, volWeight);

      const surcharges = [];

      // 2. 超尺寸附加费 I (33元/票)
      if ((maxSide > 120 && maxSide <= 175) || midSide > 60 || minSide > 60) {
        surcharges.push({ name: "超尺寸附加费 I", fee: FREIGHT_FEES.DPD.oversize });
      }

      // 3. 超体积附加费 (33元/票)
      if (volume >= 0.15) {
        surcharges.push({ name: "超体积附加费", fee: FREIGHT_FEES.DPD.volume });
      }

      // 4. 超重附加费 (282.1元/票)
      if (totalWeight >= 31.5 && totalWeight <= 40) {
        surcharges.push({ name: "超重附加费", fee: FREIGHT_FEES.DPD.overweight });
      }

      // 5. 超尺寸附加费 II (279.2元/票)
      if ((maxSide > 175 && maxSide <= 250) || girth >= 300) {
        surcharges.push({ name: "超尺寸附加费 II", fee: FREIGHT_FEES.DPD.oversizeLevel2 });
      }

      // 6. 超限附加费 (488.9元/票)
      if (maxSide > 250 || girth > 330 || totalWeight > 40) {
        surcharges.push({ name: "超限附加费", fee: FREIGHT_FEES.DPD.extreme });
      }

      return {
        isValid,
        volWeight,
        chargeWeight,
        surcharges,
        rules: {
          maxSide: maxSide <= 120,
          midSide: midSide <= 60,
          minSide: minSide <= 60,
          minLimit: maxSide >= 17 && midSide >= 13 && minSide >= 3,
          weight: totalWeight <= 31.5,
          girth: girth <= 300,
          volume: volume < 0.15,
        },
      };
    };

    return {
      baseMax,
      baseMid,
      baseMin,
      maxSide,
      midSide,
      minSide,
      totalWeight,
      girth,
      volume,
      dhl: checkDhl(),
      gls: checkGls(),
      dpd: checkDpd(),
    };
  }, [inputs, stackAxis]);

  const recommendMaxQty = useMemo(() => {
    let maxDhl = 0;
    let maxGls = 0;
    let maxDpd = 0;
    const safeL = Number(inputs.length) || 0;
    const safeW = Number(inputs.width) || 0;
    const safeH = Number(inputs.height) || 0;
    const safeWt = Number(inputs.weight) || 0;

    if (!safeL || !safeW || !safeH || !safeWt)
      return { dhl: 0, gls: 0, dpd: 0 };

    const baseDims = [safeL, safeW, safeH].sort((a, b) => b - a);
    const bMax = baseDims[0];
    const bMid = baseDims[1];
    const bMin = baseDims[2];

    const check = (carrier: keyof typeof SHIPPING_RULES, q: number) => {
      const r = SHIPPING_RULES[carrier];
      const tw = safeWt * q;

      // Try stacking on each axis
      const tryStack = (stack: number[]) => {
        const d = [...stack].sort((a, b) => b - a);
        const g = d[0] + 2 * (d[1] + d[2]);
        const v = (d[0] * d[1] * d[2]) / 1000000;
        return (
          d[0] <= r.maxSide &&
          d[1] <= r.maxMid &&
          d[2] <= r.maxMin &&
          d[0] >= r.minSide &&
          d[1] >= r.minMid &&
          d[2] >= r.minMin &&
          tw <= r.maxWeight &&
          g <= r.maxGirth &&
          (r.maxVolume ? v <= r.maxVolume : true)
        );
      };

      return (
        tryStack([bMax * q, bMid, bMin]) ||
        tryStack([bMax, bMid * q, bMin]) ||
        tryStack([bMax, bMid, bMin * q])
      );
    };

    for (let q = 1; q <= 500; q++) {
      if (check("DHL", q)) maxDhl = q;
      if (check("GLS", q)) maxGls = q;
      if (check("DPD", q)) maxDpd = q;
    }
    return { dhl: maxDhl, gls: maxGls, dpd: maxDpd };
  }, [inputs.length, inputs.width, inputs.height, inputs.weight]);

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800">
      <header className="bg-indigo-900 shadow-xl border-b border-indigo-800">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="py-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3 text-white">
              <div className="bg-gradient-to-br from-indigo-500 to-indigo-600 p-2.5 rounded-xl shadow-inner border border-indigo-400/50">
                <Database className="w-6 h-6 text-indigo-50" />
              </div>
              <div>
                <h1 className="text-xl font-black tracking-tight text-white">
                  改单助手
                </h1>
                <p className="text-indigo-200 text-xs mt-0.5 font-medium tracking-wide">
                  渠道自动寻优 / 产品资料入库 / 订单库龄路由
                </p>
              </div>
            </div>

            <nav className="flex p-1.5 bg-indigo-950/50 rounded-xl border border-indigo-800/50 items-center justify-between">
              <div className="flex gap-2">
                <button
                  onClick={() => setActiveTab("query")}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "query" ? "bg-indigo-500 text-white shadow-md" : "text-indigo-300 hover:text-white hover:bg-indigo-800/50"}`}
                >
                  <TableProperties className="w-4 h-4" /> 资费解析
                </button>
                <button
                  onClick={() => setActiveTab("calc")}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "calc" ? "bg-indigo-500 text-white shadow-md" : "text-indigo-300 hover:text-white hover:bg-indigo-800/50"}`}
                >
                  <Calculator className="w-4 h-4" /> 包装沙盘
                </button>
              </div>
              <button
                onClick={handleClearCache}
                className="ml-4 flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold text-rose-300 hover:text-white hover:bg-rose-500/80 transition-all border border-rose-500/30 hover:border-rose-500"
                title="清空所有上传底表与缓存"
              >
                <XCircle className="w-3.5 h-3.5" /> 清空重置
              </button>
            </nav>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto p-4 sm:p-6 lg:p-8">
        {activeTab === "query" && (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="p-6 sm:p-8">
              <div className="mb-6">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-4">
                  <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2 uppercase tracking-wide">
                    <span className="bg-indigo-100 text-indigo-700 w-6 h-6 flex items-center justify-center rounded-full text-xs">
                      1
                    </span>
                    加载基础数据
                  </h2>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {/* Card 1: 尾程运费表 */}
                  <div
                    className={`border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer relative group ${
                      csvData
                        ? "border-emerald-400 bg-gradient-to-br from-emerald-50 to-teal-50/50 shadow-sm"
                        : "border-slate-300 hover:border-indigo-400 bg-slate-50 hover:bg-indigo-50/30"
                    } overflow-hidden`}
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <div className="absolute top-0 left-0 bg-slate-800 text-white text-[10px] px-2 py-1 rounded-br-lg rounded-tl-xl font-bold uppercase tracking-wider z-10">
                      Required 必需
                    </div>

                    <svg
                      className="absolute -bottom-4 right-0 w-24 h-24 text-indigo-500/5 opacity-50 transform group-hover:scale-110 transition-transform duration-500 pointer-events-none"
                      viewBox="0 0 100 100"
                      fill="currentColor"
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <polygon points="50,0 100,25 100,75 50,100 0,75 0,25" />
                    </svg>

                    <input
                      type="file"
                      accept=".csv, .xlsx, .xls"
                      className="hidden"
                      ref={fileInputRef}
                      onChange={(e) => handleFileUpload(e, "freight")}
                    />

                    {loading ? (
                      <div className="flex flex-col items-center py-2 relative z-10">
                        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600 mb-2"></div>
                        <p className="text-sm text-slate-500">
                          解析尾程网点...
                        </p>
                      </div>
                    ) : csvData ? (
                      <div className="flex flex-col items-center py-2 text-emerald-800 relative z-10">
                        <CheckCircle2 className="w-8 h-8 mb-2 text-emerald-500" />
                        <p className="font-extrabold">尾程运费表已就绪</p>
                        <p
                          className="text-xs mt-1 text-emerald-600/80 truncate max-w-[200px]"
                          title={fileName}
                        >
                          {fileName}
                        </p>
                        <p className="text-xs mt-2 font-bold bg-emerald-200/50 px-3 py-1 rounded-full shadow-sm text-emerald-800">
                          {Object.keys(meta.warehouses).length} 个发货仓方案
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center py-2 text-slate-500 hover:text-indigo-600 transition-colors relative z-10">
                        <TableProperties className="w-10 h-10 mb-3 text-slate-400 group-hover:text-indigo-500 transition-colors" />
                        <p className="font-bold text-slate-700 mb-1 group-hover:text-indigo-700">
                          导入尾程报价表
                        </p>
                        <p className="text-xs text-slate-500 group-hover:text-indigo-600/70">
                          支持 Excel / CSV 利润表结构
                        </p>
                      </div>
                    )}
                  </div>

                  {/* Card 2: 优先与库存规则 */}
                  <div
                    className={`border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer relative group ${
                      Object.keys(priorityMap).length > 0
                        ? "border-sky-400 bg-gradient-to-br from-sky-50 to-blue-50/50 shadow-sm"
                        : "border-slate-300 hover:border-sky-400 bg-slate-50 hover:bg-sky-50/30"
                    } overflow-hidden`}
                    onClick={() => priorityInputRef.current?.click()}
                  >
                    <div className="absolute top-0 left-0 bg-slate-200 text-slate-600 text-[10px] px-2 py-1 rounded-br-lg rounded-tl-xl font-bold uppercase tracking-wider z-10">
                      Optional 选填
                    </div>

                    <svg
                      className="absolute -bottom-4 right-0 w-24 h-24 text-sky-500/5 opacity-50 transform group-hover:scale-110 transition-transform duration-500 pointer-events-none"
                      viewBox="0 0 100 100"
                      fill="currentColor"
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <rect x="10" y="10" width="80" height="80" rx="20" />
                    </svg>

                    <input
                      type="file"
                      accept=".csv, .xlsx, .xls"
                      className="hidden"
                      ref={priorityInputRef}
                      onChange={(e) => handleFileUpload(e, "priority")}
                    />

                    {priorityLoading ? (
                      <div className="flex flex-col items-center py-2 relative z-10">
                        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-sky-600 mb-2"></div>
                        <p className="text-sm text-slate-500">
                          解析优先与库存规则...
                        </p>
                      </div>
                    ) : Object.keys(priorityMap).length > 0 ? (
                      <div className="flex flex-col items-center py-2 text-sky-800 relative z-10">
                        <CheckSquare className="w-8 h-8 mb-2 text-sky-500" />
                        <p className="font-extrabold">规则与库存校验激活</p>
                        <p
                          className="text-xs mt-1 text-sky-600/80 truncate max-w-[200px]"
                          title={priorityFile || ""}
                        >
                          {priorityFile}
                        </p>
                        <p className="text-xs mt-2 font-bold bg-sky-200/50 px-3 py-1 rounded-full shadow-sm text-sky-800">
                          库存校验SKU: {Object.keys(inventoryMap).length}
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center py-2 text-slate-500 hover:text-sky-600 transition-colors relative z-10">
                        <Database className="w-10 h-10 mb-3 text-slate-400 group-hover:text-sky-500 transition-colors" />
                        <p className="font-bold text-slate-700 mb-1 group-hover:text-sky-700">
                          导入发货与库存文件
                        </p>
                        <p className="text-xs text-slate-500 group-hover:text-sky-600/70">
                          支持读取真实【可用库存】与优先规则
                        </p>
                      </div>
                    )}
                  </div>

                  {/* Card 3: 产品资料数据 */}
                  <div
                    className={`border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer relative group ${
                      Object.keys(productInfoMap).length > 0
                        ? "border-amber-400 bg-gradient-to-br from-amber-50 to-orange-50/50 shadow-sm"
                        : "border-slate-300 hover:border-amber-400 bg-slate-50 hover:bg-amber-50/30"
                    } overflow-hidden`}
                    onClick={() => productInputRef.current?.click()}
                  >
                    <div className="absolute top-0 left-0 bg-slate-200 text-slate-600 text-[10px] px-2 py-1 rounded-br-lg rounded-tl-xl font-bold uppercase tracking-wider z-10">
                      Optional 选填
                    </div>

                    <button
                      className="absolute top-2 right-2 text-slate-400 hover:text-amber-600 transition-colors z-20 flex items-center gap-1 text-[10px] font-bold bg-white/50 px-2 py-1 rounded-full shadow-sm"
                      onClick={downloadProductTemplate}
                      title="下载导入模板"
                    >
                      <FileDown className="w-3 h-3" />
                      下载模板
                    </button>

                    {/* SVG Background Decoration */}
                    <svg
                      className="absolute -bottom-4 right-0 w-24 h-24 text-amber-500/5 opacity-50 transform group-hover:scale-110 transition-transform duration-500 pointer-events-none"
                      viewBox="0 0 100 100"
                      fill="currentColor"
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <path d="M50 0 C77.6 0 100 22.4 100 50 C100 77.6 77.6 100 50 100 C22.4 100 0 77.6 0 50 C0 22.4 22.4 0 50 0 Z" />
                    </svg>

                    <input
                      type="file"
                      accept=".csv, .xlsx, .xls"
                      className="hidden"
                      ref={productInputRef}
                      onChange={(e) => handleFileUpload(e, "product")}
                    />

                    {productLoading ? (
                      <div className="flex flex-col items-center py-2 relative z-10">
                        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-amber-600 mb-2"></div>
                        <p className="text-sm text-slate-500">
                          解析产品数据...
                        </p>
                      </div>
                    ) : Object.keys(productInfoMap).length > 0 ? (
                      <div className="flex flex-col items-center py-2 text-amber-800 relative z-10">
                        <CheckSquare className="w-8 h-8 mb-2 text-amber-500" />
                        <p className="font-extrabold">产品箱规库已激活</p>
                        <p
                          className="text-xs mt-1 text-amber-600/80 truncate max-w-[200px]"
                          title={productFile || ""}
                        >
                          {productFile}
                        </p>
                        <p className="text-xs font-bold mt-2 bg-amber-200/50 px-3 py-1 rounded-full shadow-sm text-amber-800">
                          已配置识别: {Object.keys(productInfoMap).length} 个SKU
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center py-2 text-slate-500 hover:text-amber-600 transition-colors relative z-10">
                        <Package className="w-10 h-10 mb-3 text-slate-400 group-hover:text-amber-500 transition-colors" />
                        <p className="font-bold text-slate-700 mb-1 group-hover:text-amber-700">
                          导入产品箱规表
                        </p>
                        <p className="text-xs text-slate-500 group-hover:text-amber-600/70">
                          支持识别 SKU / 一箱件数
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {fileType === "excel" && sheetNames.length > 0 && (
                <div className="mb-8 p-5 bg-indigo-50/50 border border-indigo-100 rounded-xl flex items-center gap-4">
                  <div className="bg-indigo-100 p-2 rounded-lg">
                    <Layers className="w-5 h-5 text-indigo-600 flex-shrink-0" />
                  </div>
                  <div className="flex-1">
                    <label className="block text-xs font-bold text-indigo-900 uppercase tracking-widest mb-1.5">
                      锁定目标工作表
                    </label>
                    <select
                      value={currentSheet}
                      onChange={handleSheetChange}
                      className="w-full bg-white border border-indigo-200 text-slate-800 font-medium text-sm rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 block p-2.5 outline-none shadow-sm transition-all cursor-pointer"
                    >
                      {sheetNames.map((sheet, idx) => (
                        <option key={idx} value={sheet}>
                          {sheet}{" "}
                          {sheet.includes("尾程") ? "✨ (系统推荐)" : ""}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              )}

              {error && (
                <div className="mb-6 bg-rose-50 border border-rose-200 rounded-xl p-4 flex items-start gap-3 text-rose-800 shadow-sm">
                  <FileWarning className="w-5 h-5 flex-shrink-0 mt-0.5 text-rose-600" />
                  <div>
                    <h4 className="font-bold text-sm mb-0.5">执行中断</h4>
                    <p className="text-sm opacity-90">{error}</p>
                  </div>
                </div>
              )}

              <div
                className={`transition-opacity duration-500 ${!csvData ? "opacity-40 pointer-events-none grayscale" : "opacity-100"} mt-8`}
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-4">
                  <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2 uppercase tracking-wide">
                    <span className="bg-indigo-100 text-indigo-700 w-6 h-6 flex items-center justify-center rounded-full text-xs">
                      2
                    </span>
                    检索参数下发
                  </h2>
                  <div className="flex items-center space-x-2">
                    <input
                      id="eu-account-toggle"
                      type="checkbox"
                      checked={isEUAccount}
                      onChange={(e) => setIsEUAccount(e.target.checked)}
                      className="w-4 h-4 text-indigo-600 bg-slate-100 border-slate-300 rounded focus:ring-indigo-500 cursor-pointer"
                    />
                    <label
                      htmlFor="eu-account-toggle"
                      className="text-sm font-semibold text-slate-700 flex items-center cursor-pointer"
                    >
                      <Globe className="w-4 h-4 text-slate-400 mr-1.5" />{" "}
                      此订单属于【欧盟账号】
                    </label>
                  </div>
                </div>
                <form
                  onSubmit={handleSearch}
                  className="flex flex-col sm:flex-row gap-4 items-end"
                >
                  <div className="flex-1 w-full">
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                      <Package className="w-3.5 h-3.5" /> 目标 SKU
                    </label>
                    <div className="relative">
                      <input
                        type="text"
                        list="sku-options"
                        value={inputSku}
                        onChange={(e) => setInputSku(e.target.value)}
                        placeholder="如: EB-B-1-SG4978"
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-shadow text-slate-900 font-medium font-mono text-sm shadow-sm"
                      />
                      {inputSku && (
                        <button
                          type="button"
                          onClick={() => setInputSku("")}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors"
                        >
                          <XCircle className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                    <datalist id="sku-options">
                      {meta.skuList.map((sku, idx) => (
                        <option key={idx} value={sku} />
                      ))}
                    </datalist>
                  </div>
                  <div className="flex-1 w-full">
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                      <MapPin className="w-3.5 h-3.5" /> 目的国家
                    </label>
                    <div className="relative">
                      <input
                        type="text"
                        list="country-options"
                        value={inputCountry}
                        onChange={(e) => setInputCountry(e.target.value)}
                        placeholder="如: 爱尔兰"
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-shadow text-slate-900 font-medium text-sm shadow-sm"
                      />
                      {inputCountry && (
                        <button
                          type="button"
                          onClick={() => setInputCountry("")}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors"
                        >
                          <XCircle className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                    <datalist id="country-options">
                      {meta.countryList.map((country, idx) => (
                        <option key={idx} value={country} />
                      ))}
                    </datalist>
                  </div>
                  <button
                    type="submit"
                    className="w-full sm:w-auto bg-slate-900 hover:bg-slate-800 text-white font-bold py-3.5 px-8 rounded-xl flex items-center justify-center gap-2 transition-all shadow-md hover:shadow-lg active:scale-[0.99] text-sm"
                  >
                    <Search className="w-5 h-5" /> 智能寻源
                  </button>
                </form>
              </div>

              {/* ====== 结果展示区 ====== */}
              {result && (
                <div className="animate-in fade-in slide-in-from-bottom-4 space-y-6 mt-8">
                  <div className="flex items-center gap-3 px-2">
                    <div className="h-px bg-slate-200 flex-1"></div>
                    <span className="text-xs font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2">
                      调度中心报告
                    </span>
                    <div className="h-px bg-slate-200 flex-1"></div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                    {/* ====== 首选方案卡片 ====== */}
                    <div
                      className={`lg:col-span-8 bg-white rounded-2xl shadow-sm border overflow-hidden relative ${
                        result.recommendType === "eu-blocked" ||
                        result.recommendType === "no-stock"
                          ? "border-rose-300"
                          : "border-slate-200"
                      }`}
                    >
                      <div
                        className={`absolute top-0 right-0 w-32 h-32 rounded-bl-full -mr-16 -mt-16 opacity-10 ${
                          result.recommendType === "eu-blocked" ||
                          result.recommendType === "no-stock"
                            ? "bg-rose-500"
                            : result.recommendType?.includes("eu")
                              ? "bg-amber-500"
                              : result.recommendType === "priority"
                                ? "bg-indigo-600"
                                : "bg-emerald-500"
                        }`}
                      ></div>

                      <div className="p-8">
                        <div className="flex items-center justify-between mb-6 relative z-10">
                          <div className="flex items-center gap-2">
                            {result.recommendType === "no-stock" ? (
                              <span className="inline-flex items-center gap-1.5 bg-rose-100 text-rose-700 px-3 py-1 rounded-full text-xs font-bold border border-rose-200 shadow-sm">
                                <ShieldAlert className="w-3.5 h-3.5" /> 缺货拦截
                              </span>
                            ) : result.recommendType === "eu-blocked" ? (
                              <span className="inline-flex items-center gap-1.5 bg-rose-100 text-rose-700 px-3 py-1 rounded-full text-xs font-bold border border-rose-200 shadow-sm">
                                <ShieldAlert className="w-3.5 h-3.5" /> 严重拦截
                              </span>
                            ) : result.recommendType?.includes("eu") ? (
                              <span className="inline-flex items-center gap-1.5 bg-amber-100 text-amber-700 px-3 py-1 rounded-full text-xs font-bold border border-amber-200 shadow-sm">
                                <Globe className="w-3.5 h-3.5" /> 欧盟账号控制
                              </span>
                            ) : result.recommendType === "priority" ? (
                              <span className="inline-flex items-center gap-1.5 bg-indigo-100 text-indigo-700 px-3 py-1 rounded-full text-xs font-bold border border-indigo-200 shadow-sm">
                                <Layers className="w-3.5 h-3.5" /> 优先规则干预
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1.5 bg-emerald-100 text-emerald-700 px-3 py-1 rounded-full text-xs font-bold border border-emerald-200 shadow-sm">
                                <TrendingDown className="w-3.5 h-3.5" />{" "}
                                最低成本寻优
                              </span>
                            )}
                            <span
                              className={`text-xs font-bold ${result.recommendType === "eu-blocked" || result.recommendType === "no-stock" ? "text-rose-600" : "text-slate-500"}`}
                            >
                              {result.recommendReason}
                            </span>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 relative z-10">
                          <div>
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                              发货指令 ➔ 目标仓库
                            </p>
                            <h3
                              className={`text-4xl font-black mb-6 ${result.recommendType === "eu-blocked" || result.recommendType === "no-stock" ? "text-rose-600" : "text-slate-900"}`}
                            >
                              {result.bestOption?.warehouse || result.warehouse}
                            </h3>

                            <div className="space-y-4">
                              <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                                <span className="text-slate-500 text-sm">
                                  库龄可用库存
                                </span>
                                <div className="flex flex-wrap items-center justify-end gap-1.5 font-bold text-slate-800 flex-1 ml-4 shadow-sm-border">
                                  {result.bestOption?.stock !== undefined &&
                                  result.bestOption?.stock !== "未知" ? (
                                    <>
                                      {result.bestOption.stock.eb > 0 && (
                                        <span className="text-blue-600 bg-blue-50 border border-blue-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          EB: {result.bestOption.stock.eb}
                                        </span>
                                      )}
                                      {result.bestOption.stock.am > 0 && (
                                        <span className="text-amber-600 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          AM: {result.bestOption.stock.am}
                                        </span>
                                      )}
                                      {result.bestOption.stock.t1 > 0 && (
                                        <span className="text-violet-600 bg-violet-50 border border-violet-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          T1: {result.bestOption.stock.t1}
                                        </span>
                                      )}
                                      {result.bestOption.stock.t2 > 0 && (
                                        <span className="text-cyan-600 bg-cyan-50 border border-cyan-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          T2: {result.bestOption.stock.t2}
                                        </span>
                                      )}
                                      {result.bestOption.stock.raw > 0 && (
                                        <span className="text-emerald-600 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          普通: {result.bestOption.stock.raw}
                                        </span>
                                      )}
                                      {result.bestOption.stock.suggestedSku && (
                                        <span className="text-slate-700 bg-white border border-slate-200 px-2 py-0.5 rounded shadow-sm text-xs font-mono">
                                          发货SKU: {result.bestOption.stock.suggestedSku}
                                        </span>
                                      )}
                                      {result.bestOption.stock.total === 0 && (
                                        <span className="text-rose-600 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded shadow-sm text-xs">
                                          0 件
                                        </span>
                                      )}
                                    </>
                                  ) : (
                                    "未配置数据"
                                  )}
                                </div>
                              </div>
                              <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                                <span className="text-slate-500 text-sm">
                                  目的国家
                                </span>
                                <span className="font-bold text-slate-800">
                                  {result.country}
                                </span>
                              </div>
                            </div>
                          </div>

                          <div
                            className={`rounded-xl p-6 border flex flex-col justify-center transition-all ${
                              result.recommendType === "eu-blocked" ||
                              result.recommendType === "no-stock"
                                ? "bg-rose-50 border-rose-200 cursor-not-allowed"
                                : "bg-slate-50 border-slate-200/60 hover:border-indigo-300 hover:bg-indigo-50/20 cursor-pointer shadow-sm hover:shadow-md"
                            }`}
                            onClick={() => {
                              if (
                                result.recommendType !== "eu-blocked" &&
                                result.recommendType !== "no-stock"
                              )
                                setShowBestOptionDetails(
                                  !showBestOptionDetails,
                                );
                            }}
                          >
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2 text-center">
                              系统优选渠道
                            </p>
                            <p
                              className={`text-xl font-bold text-center mb-4 break-words px-2 ${result.recommendType === "eu-blocked" || result.recommendType === "no-stock" ? "text-rose-700" : "text-indigo-700"}`}
                            >
                              {result.bestOption?.channel || result.channel}
                            </p>

                            <div className="flex flex-col items-center">
                              <p className="text-xs font-medium text-slate-500 mb-1">
                                参考运费
                              </p>
                              <p className="text-4xl font-black text-slate-900 flex items-baseline gap-1">
                                {result.bestOption?.costNum === Infinity ||
                                result.cost === "N/A" ||
                                !result.bestOption
                                  ? "N/A"
                                  : `€${result.bestOption.costNum.toFixed(2)}`}
                              </p>
                            </div>

                            {result.recommendType !== "eu-blocked" &&
                              result.recommendType !== "no-stock" && (
                                <div className="mt-5 text-xs text-indigo-500 text-center flex items-center justify-center gap-1 font-medium bg-white/50 py-1.5 rounded-md">
                                  <Layers className="w-3.5 h-3.5" />{" "}
                                  点击查看本仓全部费用排行{" "}
                                  {showBestOptionDetails ? (
                                    <svg
                                      className="w-3.5 h-3.5"
                                      fill="none"
                                      stroke="currentColor"
                                      viewBox="0 0 24 24"
                                    >
                                      <path
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                        strokeWidth="2"
                                        d="M5 15l7-7 7 7"
                                      ></path>
                                    </svg>
                                  ) : (
                                    <svg
                                      className="w-3.5 h-3.5"
                                      fill="none"
                                      stroke="currentColor"
                                      viewBox="0 0 24 24"
                                    >
                                      <path
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                        strokeWidth="2"
                                        d="M19 9l-7 7-7-7"
                                      ></path>
                                    </svg>
                                  )}
                                </div>
                              )}
                          </div>
                        </div>

                        {showBestOptionDetails &&
                          result.bestOption?.details &&
                          result.bestOption.details.length > 0 && (
                            <div className="mt-6 pt-5 border-t border-slate-100 animate-in fade-in relative z-10">
                              <div className="flex items-center justify-between mb-3">
                                <p className="text-xs font-bold text-slate-500 uppercase">
                                  当前仓库所有可用渠道明细 (完整列表)
                                </p>
                                <span className="text-[10px] bg-slate-100 text-slate-500 px-2 py-0.5 rounded">
                                  已按费用递增排序
                                </span>
                              </div>
                              <div className="bg-slate-50 rounded-xl border border-slate-100 p-3 max-h-64 overflow-y-auto space-y-1.5 shadow-inner">
                                {result.bestOption.details.map(
                                  (detail: any, idx: number) => (
                                    <div
                                      key={idx}
                                      className="flex justify-between items-center py-2 px-4 hover:border-indigo-200 bg-white border border-slate-100 rounded-lg transition-colors gap-4 shadow-sm hover:shadow"
                                    >
                                      <span className="text-[13px] font-medium text-slate-700 flex-1 whitespace-normal break-words leading-tight">
                                        {detail.channel}
                                      </span>
                                      <span
                                        className={`text-[15px] font-bold whitespace-nowrap ${idx === 0 ? "text-emerald-600" : "text-slate-800"}`}
                                      >
                                        €{detail.costNum.toFixed(2)}
                                      </span>
                                    </div>
                                  ),
                                )}
                              </div>
                            </div>
                          )}
                      </div>
                    </div>

                    {/* ====== 侧边信息：尺寸与备选 ====== */}
                    <div className="lg:col-span-4 space-y-6">
                      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
                        <h4 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-2">
                          <Package className="w-4 h-4 text-slate-400" />{" "}
                          建档物理参数
                        </h4>
                        <div className="space-y-3">
                          <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                            <span className="text-sm text-slate-500">
                              外箱尺寸
                            </span>
                            <span
                              className={`font-mono text-sm font-semibold ${result.hasFullDims ? "text-slate-800" : "text-slate-400 italic"}`}
                            >
                              {result.dims}
                            </span>
                          </div>
                          <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                            <span className="text-sm text-slate-500">
                              总发毛重
                            </span>
                            <span
                              className={`font-mono text-sm font-semibold ${result.hasFullDims ? "text-slate-800" : "text-slate-400 italic"}`}
                            >
                              {result.weightVal}
                            </span>
                          </div>

                          <div
                            className={`flex items-center justify-between p-3 rounded-xl mt-3 ${result.hasQty ? "bg-indigo-50/70 border border-indigo-100/50" : "bg-slate-50 border border-slate-100"}`}
                          >
                            <span
                              className={`text-sm font-medium flex items-center gap-1.5 ${result.hasQty ? "text-indigo-700" : "text-slate-500"}`}
                            >
                              <Layers
                                className={`w-4 h-4 ${result.hasQty ? "text-indigo-500" : "text-slate-400"}`}
                              />{" "}
                              一箱装量
                            </span>
                            <span
                              className={`font-mono text-base font-bold ${result.hasQty ? "text-indigo-700" : "text-slate-400 italic"}`}
                            >
                              {result.hasQty
                                ? `${result.qtyVal} 件/箱`
                                : "未配置 (默认 1)"}
                            </span>
                          </div>
                        </div>

                        {result.hasFullDims && (
                          <button
                            onClick={handleLinkToCalc}
                            className="mt-4 w-full bg-slate-900 hover:bg-slate-800 text-white py-2.5 rounded-lg text-xs font-bold transition-colors flex justify-center items-center gap-2 shadow-sm relative overflow-hidden group"
                          >
                            <span className="relative z-10 flex items-center gap-2">
                              <Calculator className="w-4 h-4" />{" "}
                              智能带入包装沙盘{" "}
                              <ArrowRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
                            </span>
                          </button>
                        )}
                      </div>

                      {/* 备选方案列表 */}
                      {result.altOptions && result.altOptions.length > 0 && (
                        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
                          <h4 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <Warehouse className="w-4 h-4 text-slate-400" />{" "}
                            其他可用仓源报价
                          </h4>
                          <div className="space-y-2 max-h-[500px] overflow-y-auto pr-1">
                            {result.altOptions.map((opt: any, idx: number) => {
                              const isEUBlocked =
                                isEUAccount &&
                                ["粮仓", "粮2仓", "T2仓", "DE-粮仓"].some(
                                  (tw) => opt.warehouse.includes(tw),
                                );
                              const isNoStock =
                                result.hasInventoryData &&
                                (!opt.stock ||
                                  typeof opt.stock !== "object" ||
                                  opt.stock.total <= 0);
                              const isBlocked = isEUBlocked || isNoStock;
                              const isExpanded =
                                expandedAltWH === opt.warehouse;

                              return (
                                <div
                                  key={idx}
                                  className={`border rounded-xl transition-all overflow-hidden ${isBlocked ? "bg-slate-50 border-slate-100 opacity-60 grayscale-[0.3]" : "border-slate-200 hover:border-indigo-300 hover:shadow-sm"}`}
                                >
                                  <div
                                    className={`flex items-start justify-between p-4 ${isBlocked ? "cursor-pointer hover:bg-slate-100" : "cursor-pointer hover:bg-indigo-50/20"}`}
                                    onClick={() =>
                                      setExpandedAltWH(
                                        isExpanded ? null : opt.warehouse,
                                      )
                                    }
                                  >
                                    <div className="flex-1 min-w-0 pr-3 flex flex-col justify-center">
                                      <div className="flex items-center gap-2 flex-wrap mb-1.5">
                                        <p className="font-bold text-[15px] text-slate-800 tracking-tight leading-none">
                                          {opt.warehouse}
                                        </p>
                                        {isEUBlocked && (
                                          <span className="text-[10px] text-rose-600 font-bold border border-rose-200 px-1.5 py-0.5 rounded-md bg-rose-50 shadow-sm leading-none">
                                            欧盟限制
                                          </span>
                                        )}
                                        {isNoStock && (
                                          <span className="text-[10px] text-slate-500 font-bold border border-slate-200 px-1.5 py-0.5 rounded-md bg-slate-100 shadow-sm leading-none">
                                            无货
                                          </span>
                                        )}
                                      </div>
                                      {result.hasInventoryData &&
                                        !isNoStock &&
                                        opt.stock && (
                                          <div className="flex gap-1.5 flex-wrap mt-0.5 mb-2">
                                            {opt.stock.eb > 0 && (
                                              <span className="text-[10px] text-blue-600 font-bold border border-blue-200 px-1.5 py-0.5 rounded-md bg-blue-50 shadow-sm leading-none">
                                                EB: {opt.stock.eb}
                                              </span>
                                            )}
                                            {opt.stock.am > 0 && (
                                              <span className="text-[10px] text-amber-600 font-bold border border-amber-200 px-1.5 py-0.5 rounded-md bg-amber-50 shadow-sm leading-none">
                                                AM: {opt.stock.am}
                                              </span>
                                            )}
                                            {opt.stock.t1 > 0 && (
                                              <span className="text-[10px] text-violet-600 font-bold border border-violet-200 px-1.5 py-0.5 rounded-md bg-violet-50 shadow-sm leading-none">
                                                T1: {opt.stock.t1}
                                              </span>
                                            )}
                                            {opt.stock.t2 > 0 && (
                                              <span className="text-[10px] text-cyan-600 font-bold border border-cyan-200 px-1.5 py-0.5 rounded-md bg-cyan-50 shadow-sm leading-none">
                                                T2: {opt.stock.t2}
                                              </span>
                                            )}
                                            {opt.stock.raw > 0 && (
                                              <span className="text-[10px] text-emerald-600 font-bold border border-emerald-200 px-1.5 py-0.5 rounded-md bg-emerald-50 shadow-sm leading-none">
                                                自发: {opt.stock.raw}
                                              </span>
                                            )}
                                            {opt.stock.suggestedSku && (
                                              <span className="text-[10px] text-slate-700 font-bold border border-slate-200 px-1.5 py-0.5 rounded-md bg-white shadow-sm leading-none font-mono">
                                                发货SKU: {opt.stock.suggestedSku}
                                              </span>
                                            )}
                                            {opt.stock.eb === 0 &&
                                              opt.stock.am === 0 &&
                                              opt.stock.t1 === 0 &&
                                              opt.stock.t2 === 0 &&
                                              opt.stock.raw === 0 && (
                                                <span className="text-[10px] text-rose-500 font-bold border border-rose-200 px-1.5 py-0.5 rounded-md bg-rose-50 shadow-sm leading-none">
                                                  无货
                                                </span>
                                              )}
                                          </div>
                                        )}
                                      <p className="text-[11px] text-slate-500 break-words whitespace-normal leading-tight font-medium opacity-80">
                                        {opt.channel}
                                      </p>
                                    </div>
                                    <div className="text-right flex items-center justify-center gap-3 flex-shrink-0 self-center">
                                      <p className="font-extrabold text-lg text-slate-900 tracking-tight leading-none">
                                        €{opt.costNum.toFixed(2)}
                                      </p>
                                      <div className="text-slate-400 bg-white border border-slate-200 rounded-md p-1 shadow-sm transition-transform hover:bg-slate-50">
                                        {isExpanded ? (
                                          <svg
                                            className="w-4 h-4"
                                            fill="none"
                                            stroke="currentColor"
                                            viewBox="0 0 24 24"
                                          >
                                            <path
                                              strokeLinecap="round"
                                              strokeLinejoin="round"
                                              strokeWidth="2.5"
                                              d="M5 15l7-7 7 7"
                                            ></path>
                                          </svg>
                                        ) : (
                                          <svg
                                            className="w-4 h-4"
                                            fill="none"
                                            stroke="currentColor"
                                            viewBox="0 0 24 24"
                                          >
                                            <path
                                              strokeLinecap="round"
                                              strokeLinejoin="round"
                                              strokeWidth="2.5"
                                              d="M19 9l-7 7-7-7"
                                            ></path>
                                          </svg>
                                        )}
                                      </div>
                                    </div>
                                  </div>

                                  {isExpanded && opt.details && (
                                    <div className="bg-slate-50 border-t border-slate-200 p-3 space-y-1.5 shadow-inner">
                                      {opt.details.map(
                                        (detail: any, dIdx: number) => (
                                          <div
                                            key={dIdx}
                                            className="flex justify-between items-center py-2 px-3 bg-white border border-slate-100 hover:border-indigo-200 rounded-lg transition-colors text-xs gap-3 shadow-sm hover:shadow"
                                          >
                                            <span className="font-medium text-slate-700 break-words whitespace-normal flex-1 leading-tight">
                                              {detail.channel}
                                            </span>
                                            <span
                                              className={`font-bold whitespace-nowrap text-sm ${dIdx === 0 ? "text-emerald-600" : "text-slate-800"}`}
                                            >
                                              €{detail.costNum.toFixed(2)}
                                            </span>
                                          </div>
                                        ),
                                      )}
                                      {opt.details.length === 0 && (
                                        <p className="text-xs text-center text-slate-400 py-3 font-medium">
                                          无其他渠道明细
                                        </p>
                                      )}
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>

                    <div className="lg:col-span-12">
                      <div className="bg-amber-50 border-l-4 border-amber-400 p-4 rounded-r-xl shadow-sm">
                        <div className="flex">
                          <div className="flex-shrink-0">
                            <AlertTriangle className="h-5 w-5 text-amber-500" />
                          </div>
                          <div className="ml-3">
                            <h3 className="text-sm font-bold text-amber-800">
                              重要渠道与运送限制提醒
                            </h3>
                            <div className="mt-2 text-xs text-amber-700 space-y-1">
                              <p>
                                • <strong>GC Parcel / G2G</strong>: 仅限制发往
                                eBay 和 Temu 平台的订单。
                              </p>
                              <p>
                                • <strong>渠道尺寸拦截</strong>: DPD 超过最长边
                                250cm、围长 330cm 或 40kg 时自动排除；德国发境外 DHL 超过最长边
                                150cm、围长 300cm 或 30kg 时自动排除。
                              </p>
                              <p>
                                • <strong>规则来源</strong>: {FREIGHT_RULE_SOURCE}
                                ，已同步 DHL/GLS/DPD 最新附加费与境外 DHL 尺寸限制。
                              </p>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "calc" && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between bg-white p-6 rounded-2xl shadow-sm border border-slate-200 gap-4">
              <div className="flex items-center space-x-4">
                <div className="w-12 h-12 bg-indigo-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-indigo-200 flex-shrink-0">
                  <Calculator size={24} />
                </div>
                <div>
                  <h2 className="text-lg sm:text-xl font-black text-slate-900 tracking-tight">
                    装载体积与合规探测引擎
                  </h2>
                  <p className="text-xs sm:text-sm text-slate-500 mt-1 font-medium">
                    应用高精度浮点核算，严格校验渠道标准
                  </p>
                </div>
              </div>
              <div className="hidden md:flex space-x-2">
                <div className="px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full text-xs font-bold flex items-center border border-emerald-200">
                  <CheckCircle2 className="w-3.5 h-3.5 mr-1" /> 高精度引擎运行中
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-4 space-y-6">
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                  <h3 className="text-sm font-black mb-5 flex items-center text-slate-800 uppercase tracking-wider">
                    <Settings className="w-4 h-4 mr-2 text-indigo-500" />{" "}
                    单体参数映射
                  </h3>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5">
                        单件实重 (kg)
                      </label>
                      <input
                        type="number"
                        name="weight"
                        step="0.01"
                        min="0"
                        value={inputs.weight}
                        onChange={handleCalcInputChange}
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all shadow-inner"
                      />
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5">
                          长 (cm)
                        </label>
                        <input
                          type="number"
                          name="length"
                          min="0"
                          value={inputs.length}
                          onChange={handleCalcInputChange}
                          className="w-full font-mono text-sm bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white shadow-inner"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5">
                          宽 (cm)
                        </label>
                        <input
                          type="number"
                          name="width"
                          min="0"
                          value={inputs.width}
                          onChange={handleCalcInputChange}
                          className="w-full font-mono text-sm bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white shadow-inner"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5">
                          高 (cm)
                        </label>
                        <input
                          type="number"
                          name="height"
                          min="0"
                          value={inputs.height}
                          onChange={handleCalcInputChange}
                          className="w-full font-mono text-sm bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white shadow-inner"
                        />
                      </div>
                    </div>

                    <div className="pt-5 border-t border-slate-100">
                      <label className="block text-sm font-black text-slate-900 mb-2">
                        计划打包基数 (件)
                      </label>
                      <input
                        type="number"
                        name="qty"
                        min="1"
                        step="1"
                        value={inputs.qty}
                        onChange={handleCalcInputChange}
                        className="w-full bg-indigo-50 border-2 border-indigo-200 text-indigo-900 font-black rounded-xl px-4 py-3 focus:outline-none focus:border-indigo-500 transition-all text-center text-2xl shadow-inner"
                      />
                    </div>
                  </div>
                </div>

                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                  <h3 className="text-xs font-black text-slate-400 uppercase tracking-widest mb-4">
                    空间堆叠策略定界
                  </h3>
                  <div className="space-y-4">
                    <div>
                      <select
                        value={stackAxis}
                        onChange={(e) => setStackAxis(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 text-sm font-medium rounded-xl px-4 py-3 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer shadow-sm"
                      >
                        <option value="min">
                          沿 [最短边 {calcData.baseMin}cm] 排列 (系统推荐)
                        </option>
                        <option value="mid">
                          沿 [次长边 {calcData.baseMid}cm] 排列
                        </option>
                        <option value="max">
                          沿 [最长边 {calcData.baseMax}cm] 排列 (极易超限)
                        </option>
                      </select>
                    </div>
                    <div className="p-3.5 bg-indigo-50/80 rounded-xl flex items-start space-x-2.5 text-xs text-indigo-800 border border-indigo-100">
                      <Info className="w-4 h-4 flex-shrink-0 mt-0.5 text-indigo-500" />
                      <p className="leading-relaxed font-medium">
                        引擎严格遵循 IATA 计费规范计算围度：
                        <strong>最长边 + 2×(次长边+最短边)</strong>
                        ，并自动适配各路向抛重比计算体积重。
                      </p>
                    </div>
                  </div>
                </div>

                <div className="bg-slate-900 p-6 rounded-2xl shadow-xl text-white relative overflow-hidden group border border-slate-800">
                  <div className="absolute top-0 right-0 p-4 opacity-5 group-hover:scale-110 transition-transform duration-700 group-hover:rotate-12">
                    <Maximize2 size={100} />
                  </div>
                  <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-5 flex items-center relative z-10">
                    <Scale className="w-3.5 h-3.5 mr-2 text-indigo-400" />{" "}
                    理论单箱装载极限测算
                  </h3>
                  <div className="flex justify-between items-center relative z-10 bg-white/5 p-4 rounded-xl border border-white/10 backdrop-blur-sm">
                    <div className="text-center flex-1">
                      <p className="text-[9px] text-slate-400 uppercase font-bold tracking-widest mb-1">
                        DHL
                      </p>
                      <p className="text-2xl font-black text-white">
                        {recommendMaxQty.dhl}{" "}
                        <span className="text-[10px] font-bold opacity-50 ml-0.5">
                          件
                        </span>
                      </p>
                    </div>
                    <div className="h-8 border-r border-slate-700 mx-1"></div>
                    <div className="text-center flex-1">
                      <p className="text-[9px] text-slate-400 uppercase font-bold tracking-widest mb-1">
                        GLS
                      </p>
                      <p className="text-2xl font-black text-white">
                        {recommendMaxQty.gls}{" "}
                        <span className="text-[10px] font-bold opacity-50 ml-0.5">
                          件
                        </span>
                      </p>
                    </div>
                    <div className="h-8 border-r border-slate-700 mx-1"></div>
                    <div className="text-center flex-1">
                      <p className="text-[9px] text-slate-400 uppercase font-bold tracking-widest mb-1">
                        DPD
                      </p>
                      <p className="text-2xl font-black text-white">
                        {recommendMaxQty.dpd}{" "}
                        <span className="text-[10px] font-bold opacity-50 ml-0.5">
                          件
                        </span>
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              <div className="lg:col-span-8 space-y-6">
                <div className="bg-white p-6 sm:p-8 rounded-2xl shadow-sm border border-slate-200 grid grid-cols-1 md:grid-cols-3 gap-6 sm:gap-8">
                  <div className="flex flex-col justify-center text-center md:text-left">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 flex items-center justify-center md:justify-start gap-1">
                      <Package className="w-3 h-3" /> 成箱物理尺寸
                    </p>
                    <div className="text-xl sm:text-2xl font-black text-slate-800 flex items-baseline justify-center md:justify-start font-mono tracking-tight">
                      {calcData.maxSide}
                      <span className="text-slate-300 text-sm mx-1 font-normal">
                        ×
                      </span>
                      {calcData.midSide}
                      <span className="text-slate-300 text-sm mx-1 font-normal">
                        ×
                      </span>
                      {calcData.minSide}
                      <span className="text-slate-400 text-sm ml-1.5 font-bold">
                        cm
                      </span>
                    </div>
                  </div>

                  <div className="flex flex-col justify-center text-center md:text-left md:border-l md:border-slate-100 md:pl-8">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 flex items-center justify-center md:justify-start gap-1">
                      <Scale className="w-3 h-3" /> 成箱实重计费
                    </p>
                    <div className="text-2xl font-black text-slate-800 flex items-baseline justify-center md:justify-start font-mono">
                      {calcData.totalWeight.toFixed(2)}
                      <span className="text-slate-400 text-sm ml-1.5 font-bold">
                        kg
                      </span>
                    </div>
                  </div>

                  <div className="flex flex-col justify-center text-center md:text-left md:border-l md:border-slate-100 md:pl-8">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 flex items-center justify-center md:justify-start gap-1">
                      <Layers className="w-3 h-3" /> 周长核算 (Girth)
                    </p>
                    <div className="text-2xl font-black text-indigo-600 flex items-baseline justify-center md:justify-start font-mono">
                      {calcData.girth.toFixed(2)}
                      <span className="text-indigo-400/70 text-sm ml-1.5 font-bold">
                        cm
                      </span>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {[
                    {
                      id: "DHL",
                      data: calcData.dhl,
                      rules: SHIPPING_RULES.DHL,
                    },
                    {
                      id: "GLS",
                      data: calcData.gls,
                      rules: SHIPPING_RULES.GLS,
                    },
                    {
                      id: "DPD",
                      data: calcData.dpd,
                      rules: SHIPPING_RULES.DPD,
                    },
                  ].map(({ id, data, rules }) => (
                    <div
                      key={id}
                      className={`relative overflow-hidden rounded-xl border-2 transition-all duration-300 ${data.isValid ? "bg-white border-slate-200 shadow-sm" : "bg-slate-50 border-rose-200"}`}
                    >
                      {/* 头部：系统状态指示器 */}
                      <div
                        className={`flex items-center justify-between px-4 py-2.5 border-b ${data.isValid ? "bg-slate-50 border-slate-100" : "bg-rose-50 border-rose-100"}`}
                      >
                        <div className="flex items-center gap-2">
                          <div
                            className={`w-2 h-2 rounded-full ${data.isValid ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)] animate-pulse" : "bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]"}`}
                          />
                          <span className="text-[10px] font-black uppercase tracking-widest text-slate-900">
                            {id} 系统状态
                          </span>
                        </div>
                        <span
                          className={`text-[9px] font-mono font-black px-1.5 py-0.5 rounded border ${data.isValid ? "bg-emerald-100/50 border-emerald-200 text-emerald-700" : "bg-rose-100/50 border-rose-200 text-rose-700"}`}
                        >
                          {data.isValid ? "就绪" : "异常"}
                        </span>
                      </div>

                      <div className="p-4 space-y-4">
                        {/* 核心数据网格 */}
                        <div className="grid grid-cols-2 gap-2">
                          <div className="p-2.5 bg-slate-50 rounded border border-slate-100">
                            <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest mb-1">
                              体积重量
                            </p>
                            <p className="text-lg font-mono font-black text-slate-900 leading-none">
                              {data.volWeight}
                              <span className="text-[10px] font-bold ml-0.5 opacity-30">
                                KG
                              </span>
                            </p>
                          </div>
                          <div className="p-2.5 bg-slate-900 rounded border border-slate-800">
                            <p className="text-[8px] font-black text-indigo-300/50 uppercase tracking-widest mb-1">
                              计费重量
                            </p>
                            <p className="text-lg font-mono font-black text-indigo-400 leading-none">
                              {data.chargeWeight}
                              <span className="text-[10px] font-bold ml-0.5 opacity-30">
                                KG
                              </span>
                            </p>
                          </div>
                        </div>

                        {/* 规则列表 */}
                        <div className="bg-white rounded border border-slate-100 overflow-hidden">
                          <RuleItem
                            label="最长边"
                            value={calcData.maxSide}
                            limit={rules.maxSide}
                            unit="cm"
                            valid={data.rules.maxSide}
                          />
                          <RuleItem
                            label="次长边"
                            value={calcData.midSide}
                            limit={rules.maxMid}
                            unit="cm"
                            valid={data.rules.midSide}
                          />
                          <RuleItem
                            label="最短边"
                            value={calcData.minSide}
                            limit={rules.maxMin}
                            unit="cm"
                            valid={data.rules.minSide}
                          />
                          <RuleItem
                            label="实际重量"
                            value={calcData.totalWeight.toFixed(2)}
                            limit={rules.maxWeight}
                            unit="kg"
                            valid={data.rules.weight}
                          />
                          <RuleItem
                            label="围度(Girth)"
                            value={calcData.girth}
                            limit={rules.maxGirth}
                            unit="cm"
                            valid={data.rules.girth}
                          />
                          {(rules as any).maxVolume && (
                            <RuleItem
                              label="总体积"
                              value={calcData.volume}
                              limit={(rules as any).maxVolume}
                              unit="m³"
                              valid={data.rules.volume}
                            />
                          )}
                        </div>

                        {/* 附加费标签 */}
                        {data.surcharges && data.surcharges.length > 0 && (
                          <div className="flex flex-col gap-1 pt-2 border-t border-slate-100 mt-2">
                            <span className="text-[10px] font-bold text-slate-400 uppercase">
                              附加费预警
                            </span>
                            <div className="flex flex-wrap gap-1">
                              {data.surcharges.map((sc: any, idx: number) => (
                                <div
                                  key={idx}
                                  className="flex items-center gap-1 px-2 py-1 bg-rose-50 border border-rose-100 rounded text-[9px] font-black text-rose-700 uppercase"
                                >
                                  <AlertTriangle className="w-3 h-3" />{" "}
                                  {sc.name}: ¥{sc.fee}
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>

                {/* 决策支持面板 */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                  <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-2">
                    <Lightbulb className="w-4 h-4 text-amber-500" />{" "}
                    业务决策支持与优化建议
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-4">
                      <div className="p-4 bg-indigo-50 rounded-xl border border-indigo-100">
                        <p className="text-xs font-bold text-indigo-900 mb-2 uppercase tracking-widest">
                          最优渠道推荐
                        </p>
                        {(() => {
                          const validCarriers = [
                            { id: "DHL", data: calcData.dhl },
                            { id: "GLS", data: calcData.gls },
                            { id: "DPD", data: calcData.dpd },
                          ].filter((c) => c.data.isValid);

                          if (validCarriers.length === 0)
                            return (
                              <p className="text-sm text-rose-600 font-bold">
                                ⚠️ 当前尺寸/重量已超出所有常规渠道承运范围
                              </p>
                            );

                          // Sort by charge weight as a proxy for cost (lower is usually better)
                          const sorted = validCarriers.sort(
                            (a, b) => a.data.chargeWeight - b.data.chargeWeight,
                          );
                          return (
                            <div className="space-y-2">
                              <div className="flex items-center gap-2">
                                <span className="bg-emerald-500 text-white text-[10px] px-2 py-0.5 rounded font-black">
                                  首选
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {sorted[0].id}
                                </span>
                                <span className="text-xs text-slate-500">
                                  (计费重: {sorted[0].data.chargeWeight}kg)
                                </span>
                              </div>
                              {sorted.length > 1 && (
                                <div className="flex items-center gap-2 opacity-60">
                                  <span className="bg-slate-400 text-white text-[10px] px-2 py-0.5 rounded font-black">
                                    备选
                                  </span>
                                  <span className="text-sm font-bold text-slate-800">
                                    {sorted[1].id}
                                  </span>
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </div>
                      <div className="p-4 bg-amber-50 rounded-xl border border-amber-100">
                        <p className="text-xs font-bold text-amber-900 mb-2 uppercase tracking-widest">
                          成本优化提醒
                        </p>
                        <ul className="text-xs text-amber-800 space-y-1.5 list-disc pl-4">
                          {calcData.totalWeight > 20 && (
                            <li>当前重量较大，建议核实托盘费或超重附加费。</li>
                          )}
                          {calcData.girth > 250 && (
                            <li>
                              周长接近 300cm 临界点，务必预留 2-3cm 测量误差。
                            </li>
                          )}
                          {calcData.maxSide > 100 && (
                            <li>
                              长边超过 100cm，部分渠道可能产生人工处理费。
                            </li>
                          )}
                          {calcData.volume > 0.12 && (
                            <li>
                              体积接近 0.15m³，GLS/DPD 极易产生体积附加费。
                            </li>
                          )}
                          {calcData.totalWeight < calcData.gls.volWeight && (
                            <li>
                              GLS 抛重比为
                              6000，当前处于抛货状态，建议压缩包装。
                            </li>
                          )}
                        </ul>
                      </div>
                    </div>

                    <div className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                        <p className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-widest">
                        GLS/DPD 5.15规则摘要
                      </p>
                      <div className="space-y-3">
                        <div className="flex justify-between text-[10px] border-b border-slate-200 pb-1.5">
                          <span className="text-slate-400">GLS体积附加费阈值</span>
                          <span className="font-bold text-slate-700">
                            0.15 m³ (约 {FREIGHT_FEES.GLS.volume} RMB)
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px] border-b border-slate-200 pb-1.5">
                          <span className="text-slate-400">GLS超长附加费</span>
                          <span className="font-bold text-slate-700">
                            &gt;150 cm (约 {FREIGHT_FEES.GLS.overLength} RMB)
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px] border-b border-slate-200 pb-1.5">
                          <span className="text-slate-400">DPD超重附加费</span>
                          <span className="font-bold text-slate-700">
                            31.5-40 kg (约 {FREIGHT_FEES.DPD.overweight} RMB)
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px] border-b border-slate-200 pb-1.5">
                          <span className="text-slate-400">极度超限罚金</span>
                          <span className="font-bold text-rose-600">
                            GLS {FREIGHT_FEES.GLS.extreme} / DPD {FREIGHT_FEES.DPD.extreme} RMB
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px]">
                          <span className="text-slate-400">
                            非标准包装 (NC)
                          </span>
                          <span className="font-bold text-slate-700">
                            GLS {FREIGHT_FEES.GLS.nonConveyable} RMB (手动分拣)
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* 所有渠道报价弹窗 */}
      {showAllCosts && result && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[80vh]">
            <div className="flex items-center justify-between p-5 border-b border-slate-100 bg-slate-50">
              <h3 className="font-black text-slate-800 flex items-center gap-2">
                <Database className="w-5 h-5 text-indigo-500" />
                所有可用渠道报价
              </h3>
              <button
                onClick={() => setShowAllCosts(false)}
                className="text-slate-400 hover:text-slate-600 transition-colors"
              >
                <XCircle className="w-6 h-6" />
              </button>
            </div>
            <div className="p-5 overflow-y-auto flex-1">
              <div className="mb-4 flex items-center justify-between text-sm">
                <span className="font-bold text-slate-600">
                  SKU:{" "}
                  <span className="font-mono text-slate-900">{result.sku}</span>
                </span>
                <span className="font-bold text-slate-600">
                  流向: <span className="text-slate-900">{result.country}</span>
                </span>
              </div>
              <div className="space-y-2">
                {result.allCosts.map((c: any, idx: number) => (
                  <div
                    key={idx}
                    className="flex items-center justify-between p-3 rounded-xl border border-slate-100 bg-slate-50 hover:bg-indigo-50 hover:border-indigo-100 transition-colors"
                  >
                    <div className="flex items-center gap-3">
                      <span className="w-6 h-6 rounded-full bg-slate-200 text-slate-500 flex items-center justify-center text-xs font-black">
                        {idx + 1}
                      </span>
                      <span className="font-bold text-slate-800">
                        {c.channel}
                      </span>
                    </div>
                    <span className="font-black text-emerald-600 font-mono">
                      €{c.cost.toFixed(2)}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
