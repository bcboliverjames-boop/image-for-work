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
      
      // 插入矩形形状
      const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      
      // 设置形状属性
      shape.name = shapeName;
      shape.width = widthPts;
      shape.height = heightPts;
      shape.left = 100;
      shape.top = 100;
      
      // 设置填充样式
      shape.fill.setSolidColor(REFERENCE_BOX_CONFIG.fillColor);
      shape.fill.transparency = REFERENCE_BOX_CONFIG.fillTransparency;
      
      // 设置边框样式
      shape.lineFormat.color = REFERENCE_BOX_CONFIG.lineColor;
      shape.lineFormat.weight = REFERENCE_BOX_CONFIG.lineWeight;
      
      await context.sync();
      
      return shapeName;
    });
  } catch (e) {
    console.error("[参考框] Excel 插入失败:", e);
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
    console.error("[参考框] Excel 读取尺寸失败:", e);
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
    console.error("[参考框] Excel 移除失败:", e);
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
    console.error("[参考框] Excel 查找失败:", e);
    return null;
  }
}
