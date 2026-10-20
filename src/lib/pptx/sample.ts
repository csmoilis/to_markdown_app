import { buildPptx } from "./build.ts";

function renderBanner(): Promise<Uint8Array> {
  const canvas = document.createElement("canvas");
  canvas.width = 880;
  canvas.height = 280;
  const context = canvas.getContext("2d");
  if (!context) return Promise.reject(new Error("Could not draw the sample image."));
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#1c1c1a";
  context.font = "bold 54px sans-serif";
  context.fillText("NORTHWIND 2026", 48, 110);
  context.font = "36px sans-serif";
  context.fillText("Invoice total 18420", 48, 190);
  return new Promise((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        reject(new Error("Could not encode the sample image."));
        return;
      }
      resolve(new Uint8Array(await blob.arrayBuffer()));
    }, "image/png");
  });
}

export async function buildSampleFile(): Promise<{ name: string; buffer: ArrayBuffer }> {
  const png = await renderBanner();
  const buffer = await buildPptx([
    {
      title: "Folio sample",
      titleBold: true,
      subtitle: "Words, a table, notes, and a picture",
      bullets: ["Slide text becomes markdown", "Speaker notes stay with the slide"],
      groupText: "Inside a grouped shape",
      fallbackText: "Fallback wording",
      choiceText: "Choice wording",
      notes: "Mention that image text is read in the browser.",
    },
    {
      title: "Quarterly snapshot",
      table: [
        ["Region", "Revenue"],
        ["North", "18420"],
        ["South", "9600"],
      ],
      hidden: true,
    },
    {
      title: "Banner",
      image: { bytes: png, ext: "png" },
    },
  ]);
  return { name: "folio-sample.pptx", buffer };
}
