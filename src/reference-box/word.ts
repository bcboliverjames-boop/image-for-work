/**
 * Word 参考框实现 v2
 * 使用 OOXML 插入矩形形状，通过 altTextDescription 识别参考框
 */

import {
  REFERENCE_BOX_CONFIG,
  ReferenceBoxSize,
  cmToPoints,
  pointsToCm,
  generateReferenceBoxName,
} from "./types";

// EMU 转换：1 point = 12700 EMU
const POINTS_TO_EMU = 12700;

// 参考框的特殊描述标记，用于识别
const REF_BOX_DESCRIPTION = "OfficePasteWidth_ReferenceBox";

function pointsToEmu(points: number): number {
  return Math.round(points * POINTS_TO_EMU);
}

/**
 * 生成矩形形状的 OOXML
 */
function generateRectangleOoxml(shapeName: string, widthPts: number, heightPts: number): string {
  const widthEmu = pointsToEmu(widthPts);
  const heightEmu = pointsToEmu(heightPts);
  
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                  mc:Ignorable="wp14">
        <w:body>
          <w:p>
            <w:r>
              <mc:AlternateContent>
                <mc:Choice Requires="wps">
                  <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" 
                               relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" 
                               allowOverlap="1">
                      <wp:simplePos x="0" y="0"/>
                      <wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                      <wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>
                      <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
                      <wp:effectExtent l="0" t="0" r="19050" b="19050"/>
                      <wp:wrapNone/>
                      <wp:docPr id="1" name="${shapeName}" descr="${REF_BOX_DESCRIPTION}"/>
                      <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks noChangeAspect="0"/>
                      </wp:cNvGraphicFramePr>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                          <wps:wsp>
                            <wps:cNvSpPr txBox="0">
                              <a:spLocks noChangeArrowheads="1"/>
                            </wps:cNvSpPr>
                            <wps:spPr bwMode="auto">
                              <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
                              </a:xfrm>
                              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                              <a:solidFill>
                                <a:srgbClr val="FFCC00"><a:alpha val="15000"/></a:srgbClr>
                              </a:solidFill>
                              <a:ln w="19050">
                                <a:solidFill><a:srgbClr val="FF6600"/></a:solidFill>
                              </a:ln>
                            </wps:spPr>
                            <wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr" anchorCtr="0" upright="1">
                              <a:noAutofit/>
                            </wps:bodyPr>
                          </wps:wsp>
                        </a:graphicData>
                      </a:graphic>
                    </wp:anchor>
                  </w:drawing>
                </mc:Choice>
              </mc:AlternateContent>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * 检查形状是否为参考框（通过 altTextDescription）
 */
function isReferenceBox(shape: Word.Shape): boolean {
  try {
    // altTextDescription 包含我们的标记
    return shape.altTextDescription === REF_BOX_DESCRIPTION;
  } catch {
    return false;
  }
}

/**
 * 在 Word 文档中插入参考框
 */
export async function insertWordReferenceBox(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const shapeName = generateReferenceBoxName();
      const widthPts = cmToPoints(REFERENCE_BOX_CONFIG.defaultWidthCm);
      const heightPts = cmToPoints(REFERENCE_BOX_CONFIG.defaultHeightCm);
      
      const ooxml = generateRectangleOoxml(shapeName, widthPts, heightPts);
      
      // 在当前选择位置插入 OOXML
      const selection = context.document.getSelection();
      selection.insertOoxml(ooxml, Word.InsertLocation.after);
      
      await context.sync();
      
      console.log("[参考框] 插入成功:", shapeName);
      return shapeName;
    });
  } catch (e) {
    console.error("[参考框] Word 插入失败:", e);
    return null;
  }
}

/**
 * 获取 Word 参考框尺寸
 * 通过 altTextDescription 识别参考框
 */
export async function getWordReferenceBoxSize(shapeName: string): Promise<ReferenceBoxSize | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const shapes = body.shapes;
      shapes.load("items");
      await context.sync();
      
      for (const shape of shapes.items) {
        shape.load(["altTextDescription", "width", "height"]);
      }
      await context.sync();
      
      for (const shape of shapes.items) {
        if (isReferenceBox(shape)) {
          const result = {
            widthCm: Math.round(pointsToCm(shape.width) * 10) / 10,
            heightCm: Math.round(pointsToCm(shape.height) * 10) / 10,
          };
          console.log("[参考框] 读取尺寸:", result);
          return result;
        }
      }
      
      console.log("[参考框] 未找到参考框");
      return null;
    });
  } catch (e) {
    console.error("[参考框] Word 读取尺寸失败:", e);
    return null;
  }
}

/**
 * 移除 Word 参考框
 * 通过 altTextDescription 识别并删除所有参考框
 */
export async function removeWordReferenceBox(shapeName: string): Promise<boolean> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const shapes = body.shapes;
      shapes.load("items");
      await context.sync();
      
      console.log("[参考框] 开始删除，形状总数:", shapes.items.length);
      
      for (const shape of shapes.items) {
        shape.load(["altTextDescription", "name"]);
      }
      await context.sync();
      
      let deleted = false;
      for (const shape of shapes.items) {
        console.log("[参考框] 检查形状:", { name: shape.name, desc: shape.altTextDescription });
        if (isReferenceBox(shape)) {
          console.log("[参考框] 找到参考框，删除:", shape.name);
          shape.delete();
          deleted = true;
        }
      }
      
      if (deleted) {
        await context.sync();
        console.log("[参考框] 删除成功");
        return true;
      }
      
      console.log("[参考框] 未找到参考框");
      return false;
    });
  } catch (e) {
    console.error("[参考框] Word 移除失败:", e);
    return false;
  }
}

/**
 * 查找 Word 文档中的参考框
 */
export async function findWordReferenceBox(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const shapes = body.shapes;
      shapes.load("items");
      await context.sync();
      
      for (const shape of shapes.items) {
        shape.load(["altTextDescription", "name"]);
      }
      await context.sync();
      
      for (const shape of shapes.items) {
        if (isReferenceBox(shape)) {
          return shape.name;
        }
      }
      
      return null;
    });
  } catch (e) {
    console.error("[参考框] Word 查找失败:", e);
    return null;
  }
}
