/**
 * Excel 参考框实现
 * 使用 Excel.js API 在工作表中插入、读取、移除参考框
 */

import {
  REFERENCE_BOX_CONFIG,
  ReferenceBoxSize,
  cmToPoints,
  pointsToCm,
  generateReferenceBoxName,
  isReferenceBoxName,
} from "./types";

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[refBoxExcel]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[refBoxExcel]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[refBoxExcel]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[refBoxExcel]", msg, data ?? ""),
  };
}

function toLoggableError(e: unknown): any {
  if (e && typeof e === "object") {
    const anyErr = e as any;
    return {
      name: anyErr?.name,
      message: anyErr?.message ?? String(e),
      code: anyErr?.code,
      debugInfo: anyErr?.debugInfo,
      httpStatusCode: anyErr?.httpStatusCode,
    };
  }
  return { message: String(e) };
}

/**
 * 在 Excel 工作表中插入参考框
 * @returns 形状名称，失败返回 null
 */
export async function insertExcelReferenceBox(): Promise<string | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 生成唯一名称
      const shapeName = generateReferenceBoxName();
      
      // 计算默认尺寸（磅）
      const widthPts = cmToPoints(REFERENCE_BOX_CONFIG.defaultWidthCm);
      const heightPts = cmToPoints(REFERENCE_BOX_CONFIG.defaultHeightCm);

      // 计算插入位置：尽量放在当前选区中心（通常在当前可视区域）。
      let leftPts = 100;
      let topPts = 100;
      const marginPts = 12;
      let anchorKind: "selectedRange" | "activeCell" | "fallback" = "fallback";
      let anchorAddress: string | null = null;
      let anchorRect: { left: number; top: number; width: number; height: number } | null = null;
      try {
        const selected = context.workbook.getSelectedRange();
        selected.load(["address", "left", "top", "width", "height"]);
        await context.sync();

        anchorKind = "selectedRange";
        anchorAddress = selected.address;
        anchorRect = { left: selected.left, top: selected.top, width: selected.width, height: selected.height };

        const centerX = selected.left + selected.width / 2;
        const centerY = selected.top + selected.height / 2;

        leftPts = Math.max(0, centerX - widthPts / 2);
        topPts = Math.max(0, centerY - heightPts / 2);
      } catch (e) {
        try {
          const cell = context.workbook.getActiveCell();
          cell.load(["address", "left", "top", "width", "height"]);
          await context.sync();

          anchorKind = "activeCell";
          anchorAddress = cell.address;
          anchorRect = { left: cell.left, top: cell.top, width: cell.width, height: cell.height };

          const centerX = cell.left + cell.width / 2;
          const centerY = cell.top + cell.height / 2;

          leftPts = Math.max(0, centerX - widthPts / 2);
          topPts = Math.max(0, centerY - heightPts / 2);
        } catch (e2) {
          getLog().warn("insertExcelReferenceBox: failed to compute position", {
            selectedRange: toLoggableError(e),
            activeCell: toLoggableError(e2),
          });
        }
      }

      getLog().info("insertExcelReferenceBox: computed position", {
        anchorKind,
        anchorAddress,
        anchorRect,
        leftPts,
        topPts,
        widthPts,
        heightPts,
      });
      
      // 插入矩形形状
      const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      
      // 设置形状属性
      shape.name = shapeName;
      shape.width = widthPts;
      shape.height = heightPts;
      shape.left = leftPts;
      shape.top = topPts;

      // 先确保“形状本体”创建成功，避免后续样式属性导致整个插入失败。
      await context.sync();

      // 样式尽力设置：某些 Excel 版本对 lineFormat/weight 支持不完整。
      try {
        shape.fill.setSolidColor(REFERENCE_BOX_CONFIG.fillColor);
        shape.fill.transparency = REFERENCE_BOX_CONFIG.fillTransparency;
      } catch (e) {
        getLog().warn("insertExcelReferenceBox: set fill failed", toLoggableError(e));
      }
      try {
        shape.lineFormat.color = REFERENCE_BOX_CONFIG.lineColor;
      } catch (e) {
        getLog().warn("insertExcelReferenceBox: set line color failed", toLoggableError(e));
      }
      try {
        const safeWeight = Math.max(1, Math.round(Number(REFERENCE_BOX_CONFIG.lineWeight) || 1));
        shape.lineFormat.weight = safeWeight;
      } catch (e) {
        getLog().warn("insertExcelReferenceBox: set line weight failed", toLoggableError(e));
      }
      try {
        await context.sync();
      } catch (e) {
        getLog().warn("insertExcelReferenceBox: style sync failed", toLoggableError(e));
      }

      getLog().info("insertExcelReferenceBox: inserted", { shapeName });

      return shapeName;
    });
  } catch (e) {
    const log = getLog();
    log.error("insertExcelReferenceBox failed", toLoggableError(e));
    return null;
  }
}

/**
 * 获取 Excel 参考框尺寸
 * @param shapeName 形状名称
 * @returns 尺寸对象，失败返回 null
 */
export async function getExcelReferenceBoxSize(shapeName: string): Promise<ReferenceBoxSize | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();
      
      // 查找指定名称的形状
      for (const shape of shapes.items) {
        shape.load(["name", "width", "height"]);
      }
      await context.sync();
      
      for (const shape of shapes.items) {
        if (shape.name === shapeName) {
          return {
            widthCm: Math.round(pointsToCm(shape.width) * 10) / 10,
            heightCm: Math.round(pointsToCm(shape.height) * 10) / 10,
          };
        }
      }
      
      return null;
    });
  } catch (e) {
    getLog().error("getExcelReferenceBoxSize failed", toLoggableError(e));
    return null;
  }
}

/**
 * 移除 Excel 参考框
 * @param shapeName 形状名称
 * @returns 是否成功
 */
export async function removeExcelReferenceBox(shapeName: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();
      
      // 查找并删除指定名称的形状
      for (const shape of shapes.items) {
        shape.load("name");
      }
      await context.sync();
      
      for (const shape of shapes.items) {
        if (shape.name === shapeName) {
          shape.delete();
          await context.sync();
          return true;
        }
      }
      
      return false;
    });
  } catch (e) {
    getLog().error("removeExcelReferenceBox failed", toLoggableError(e));
    return false;
  }
}

/**
 * 查找 Excel 工作表中的参考框
 * @returns 形状名称，不存在返回 null
 */
export async function findExcelReferenceBox(): Promise<string | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();
      
      // 遍历所有形状，查找名称前缀匹配的
      for (const shape of shapes.items) {
        shape.load("name");
      }
      await context.sync();
      
      for (const shape of shapes.items) {
        if (isReferenceBoxName(shape.name)) {
          return shape.name;
        }
      }
      
      return null;
    });
  } catch (e) {
    getLog().error("findExcelReferenceBox failed", toLoggableError(e));
    return null;
  }
}
