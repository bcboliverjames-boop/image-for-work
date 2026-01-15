/**
 * 参考框核心模块
 * 提供统一的跨应用接口，根据当前 Office 应用分发到对应实现
 */

import { ReferenceBoxState, ReferenceBoxSize, REFERENCE_BOX_CONFIG } from "./types";
import {
  insertWordReferenceBox,
  getWordReferenceBoxSize,
  removeWordReferenceBox,
  findWordReferenceBox,
} from "./word";
import {
  insertExcelReferenceBox,
  getExcelReferenceBoxSize,
  removeExcelReferenceBox,
  findExcelReferenceBox,
} from "./excel";
import {
  insertPptReferenceBox,
  getPptReferenceBoxSize,
  removePptReferenceBox,
  findPptReferenceBox,
} from "./powerpoint";

/**
 * 安全获取 Office host 类型
 */
function getOfficeHost(): typeof Office.HostType[keyof typeof Office.HostType] | null {
  try {
    return Office?.context?.host ?? null;
  } catch {
    return null;
  }
}

/**
 * 插入参考框
 * 根据当前 Office 应用自动选择对应实现
 * @returns 形状名称，失败返回 null
 */
export async function insertReferenceBox(): Promise<string | null> {
  const host = getOfficeHost();
  
  switch (host) {
    case Office.HostType.Word:
      return await insertWordReferenceBox();
    
    case Office.HostType.Excel:
      return await insertExcelReferenceBox();
    
    case Office.HostType.PowerPoint:
      return await insertPptReferenceBox();
    
    default:
      console.error("[参考框] 不支持的 Office 应用:", host);
      return null;
  }
}

/**
 * 获取参考框尺寸
 * @param shapeName 形状名称
 * @returns 尺寸对象，失败返回 null
 */
export async function getReferenceBoxSize(shapeName: string): Promise<ReferenceBoxSize | null> {
  const host = getOfficeHost();
  if (!host) return null;
  
  switch (host) {
    case Office.HostType.Word:
      return await getWordReferenceBoxSize(shapeName);
    
    case Office.HostType.Excel:
      return await getExcelReferenceBoxSize(shapeName);
    
    case Office.HostType.PowerPoint:
      return await getPptReferenceBoxSize(shapeName);
    
    default:
      return null;
  }
}

/**
 * 移除参考框
 * @param shapeName 形状名称
 * @returns 是否成功
 */
export async function removeReferenceBox(shapeName: string): Promise<boolean> {
  const host = getOfficeHost();
  if (!host) return false;
  
  switch (host) {
    case Office.HostType.Word:
      return await removeWordReferenceBox(shapeName);
    
    case Office.HostType.Excel:
      return await removeExcelReferenceBox(shapeName);
    
    case Office.HostType.PowerPoint:
      return await removePptReferenceBox(shapeName);
    
    default:
      return false;
  }
}

/**
 * 查找文档中的参考框
 * @returns 形状名称，不存在返回 null
 */
export async function findReferenceBox(): Promise<string | null> {
  const host = getOfficeHost();
  if (!host) return null;
  
  switch (host) {
    case Office.HostType.Word:
      return await findWordReferenceBox();
    
    case Office.HostType.Excel:
      return await findExcelReferenceBox();
    
    case Office.HostType.PowerPoint:
      return await findPptReferenceBox();
    
    default:
      return null;
  }
}

/**
 * 获取参考框完整状态
 * 包括是否存在、名称、当前尺寸
 * @returns 参考框状态对象
 */
export async function getReferenceBoxState(): Promise<ReferenceBoxState> {
  const shapeName = await findReferenceBox();
  
  if (!shapeName) {
    return {
      exists: false,
      shapeName: null,
      widthCm: REFERENCE_BOX_CONFIG.defaultWidthCm,
      heightCm: REFERENCE_BOX_CONFIG.defaultHeightCm,
    };
  }
  
  const size = await getReferenceBoxSize(shapeName);
  
  return {
    exists: true,
    shapeName,
    widthCm: size?.widthCm ?? REFERENCE_BOX_CONFIG.defaultWidthCm,
    heightCm: size?.heightCm ?? REFERENCE_BOX_CONFIG.defaultHeightCm,
  };
}
