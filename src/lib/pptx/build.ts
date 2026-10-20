import JSZip from "jszip";

export type BuiltParagraph = {
  text: string;
  level?: number;
  bullet?: boolean;
  bold?: boolean;
};

export type BuiltSlide = {
  title?: string;
  titleBold?: boolean;
  subtitle?: string;
  bullets?: string[];
  groupText?: string;
  fallbackText?: string;
  choiceText?: string;
  table?: string[][];
  notes?: string;
  image?: { bytes: Uint8Array; ext: "png" | "jpg" };
  chart?: boolean;
  hidden?: boolean;
};

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

function esc(value: string): string {
  return value
    .replace(/&/g, "\u0026amp;")
    .replace(/</g, "\u0026lt;")
    .replace(/>/g, "\u0026gt;")
    .replace(/"/g, "\u0026quot;");
}

function paragraphXml(paragraph: BuiltParagraph): string {
  const level = paragraph.level ?? 0;
  const bullet = paragraph.bullet
    ? `<a:buFont typeface="Arial"/><a:buChar char="•"/>`
    : `<a:buNone/>`;
  const bold = paragraph.bold ? ` b="1"` : "";
  return `<a:p><a:pPr lvl="${level}">${bullet}</a:pPr><a:r><a:rPr lang="en-US"${bold}/><a:t>${esc(paragraph.text)}</a:t></a:r></a:p>`;
}

function shapeXml(id: number, name: string, ph: string | null, x: number, y: number, paragraphs: BuiltParagraph[]): string {
  const phXml = ph ? `<p:nvPr><p:ph type="${ph}"/></p:nvPr>` : `<p:nvPr/>`;
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr/>${phXml}</p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="4000000" cy="600000"/></a:xfrm></p:spPr>
    <p:txBody><a:bodyPr/>${paragraphs.map(paragraphXml).join("")}</p:txBody>
  </p:sp>`;
}

function slideXml(slide: BuiltSlide, imageRel: string | null): string {
  const shapes: string[] = [];
  let id = 2;
  if (slide.title) {
    shapes.push(
      shapeXml(id++, "Title", "ctrTitle", 400000, 200000, [
        { text: slide.title, bold: slide.titleBold },
      ]),
    );
  }
  if (slide.subtitle) {
    shapes.push(shapeXml(id++, "Subtitle", "subTitle", 400000, 900000, [{ text: slide.subtitle }]));
  }
  if (slide.bullets?.length) {
    shapes.push(
      shapeXml(
        id++,
        "Body",
        "body",
        400000,
        1600000,
        slide.bullets.map((text) => ({ text, bullet: true })),
      ),
    );
  }
  if (slide.groupText) {
    shapes.push(`<p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="${id++}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="3200000"/><a:ext cx="4000000" cy="800000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="800000"/></a:xfrm></p:grpSpPr>
      ${shapeXml(id++, "Grouped", null, 100000, 0, [{ text: slide.groupText }])}
    </p:grpSp>`);
  }
  if (slide.fallbackText) {
    shapes.push(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="a14">${shapeXml(id++, "Choice", null, 0, 4200000, [{ text: slide.choiceText ?? "CHOICE ONLY" }])}</mc:Choice>
      <mc:Fallback>${shapeXml(id++, "Fallback", null, 0, 4200000, [{ text: slide.fallbackText }])}</mc:Fallback>
    </mc:AlternateContent>`);
  }
  if (slide.table?.length) {
    const rows = slide.table
      .map(
        (row) =>
          `<a:tr h="300000">${row
            .map(
              (cell) =>
                `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>${esc(cell)}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`,
            )
            .join("")}</a:tr>`,
      )
      .join("");
    shapes.push(`<p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="${id++}" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="400000" y="5200000"/><a:ext cx="4000000" cy="1200000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid>${slide.table[0].map(() => `<a:gridCol w="2000000"/>`).join("")}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic>
    </p:graphicFrame>`);
  }
  if (imageRel) {
    shapes.push(`<p:pic>
      <p:nvPicPr><p:cNvPr id="${id++}" name="Banner" descr="Sample banner"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="${imageRel}"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="400000" y="6800000"/><a:ext cx="3200000" cy="900000"/></a:xfrm></p:spPr>
    </p:pic>`);
  }
  if (slide.chart) {
    shapes.push(`<p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="${id++}" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="5000000" y="200000"/><a:ext cx="2000000" cy="2000000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic>
    </p:graphicFrame>`);
  }
  const show = slide.hidden ? ` show="0"` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"${show}>
  <p:cSld name="Slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${shapes.join("\n")}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function notesXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    ${shapeXml(2, "Notes", "body", 0, 0, [{ text }])}
    ${shapeXml(3, "Number", "sldNum", 0, 100, [{ text: "9" }])}
  </p:spTree></p:cSld>
</p:notes>`;
}

export async function buildPptx(slides: BuiltSlide[]): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const slideRels: string[] = [];
  slides.forEach((slide, index) => {
    const n = index + 1;
    const relId = `rId${n + 1}`;
    slideRels.push(
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/>`,
    );
    const imageRel = slide.image ? "rIdImg" : null;
    const rels = [
      slide.notes
        ? `<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${n}.xml"/>`
        : "",
      imageRel
        ? `<Relationship Id="${imageRel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${n}.${slide.image?.ext}"/>`
        : "",
    ]
      .filter(Boolean)
      .join("");
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(slide, imageRel));
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
    );
    if (slide.notes) zip.file(`ppt/notesSlides/notesSlide${n}.xml`, notesXml(slide.notes));
    if (slide.image) zip.file(`ppt/media/image${n}.${slide.image.ext}`, slide.image.bytes);
  });
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="${NS_P}" xmlns:r="${NS_R}">
  <p:sldIdLst>
    ${slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join("")}
  </p:sldIdLst>
</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slideRels.join("")}</Relationships>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
}
