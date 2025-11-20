/* eslint-disable office-addins/no-context-sync-in-loop */
/* global PowerPoint console */
import { getFormattedReferences } from "./gemini";

export async function insertText(text: string) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      const slide = slides.items[0];
      const textBox = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function analyzeDocument(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      if (!slides.items || slides.items.length === 0) {
        return "No slides found in the presentation";
      }

      // Collect all text from all shapes in selected slides
      let allText = "";
      const paragraphTexts: string[] = [];
      const shapesWithText: { shape: PowerPoint.Shape; text: string }[] = [];

      for (let slide of slides.items) {
        try {
          const shapes = slide.shapes;
          shapes.load("items");
          await context.sync();

          if (!shapes.items) continue;

          for (let shape of shapes.items) {
            if (!shape || !shape.textFrame) continue;

            try {
              const textRange = shape.textFrame.textRange;
              textRange.load("text");
              await context.sync();

              const text = textRange?.text?.trim();
              console.log({ text });
              if (text) {
                allText += text + "\n";
                paragraphTexts.push(text);
                shapesWithText.push({ shape, text });
              }
            } catch (error) {
              console.warn("Error loading shape text:", error);
              continue;
            }
          }
        } catch (error) {
          console.warn("Error processing slide:", error);
          continue;
        }
      }

      console.log("check 1");

      // If no text was found, return early
      if (!allText) {
        return "No text content found in selected slides";
      }

      console.log("check 2");

      const referenceHeaders = [
        "Reference List",
        "References List",
        "References",
        "REFERENCES LIST",
        "REFERENCE LIST",
        "REFERENCES",
        "List of References",
        "List of references",
      ];

      let referenceIndex = -1;
      for (const header of referenceHeaders) {
        const index = allText.lastIndexOf(header);
        if (index !== -1 && (referenceIndex === -1 || index > referenceIndex)) {
          referenceIndex = index;
        }
      }

      console.log("check 3");

      if (referenceIndex === -1) {
        return "No Reference List section found";
      }

      console.log("check 4");
      console.log({ allText });

      let references: string[] = [];
      try {
        console.log("referenceIndex ", referenceIndex);
        const referenceSection = allText.substring(referenceIndex);
        console.log("referenceSection ", referenceSection);

        const formattedRefs = await getFormattedReferences(referenceSection);
        references = formattedRefs.split("\n\n").map((ref) => ref.trim());

        if (references.length === 0) {
          return "No valid references found in the Reference List section";
        }
      } catch (error) {
        console.error("Error in getFormattedReferences:", error);
        throw error;
      }
      console.log("references ", references);

      console.log("check 5");

      // Find the last reference list index
      const lastReferenceListIndex = paragraphTexts
        .reverse()
        .findIndex((p) => referenceHeaders.some((header) => p.includes(header)));
      const actualIndex = lastReferenceListIndex === -1 ? -1 : paragraphTexts.length - 1 - lastReferenceListIndex;

      console.log("check 6");

      // Filter shapes to modify
      const excludeIndexes = shapesWithText
        .map((item, index) =>
          item.text.endsWith(": ") || item.text.split(" ").length <= 11 || index > actualIndex ? index : -1
        )
        .filter((index) => index !== -1);

      console.log("check 7");

      const availableIndexes = Array.from({ length: shapesWithText.length }, (_, i) => i).filter(
        (index) => !excludeIndexes.includes(index)
      );

      console.log("check 8");

      if (availableIndexes.length === 0) {
        return "No suitable text shapes found for adding references";
      }

      console.log("check 9");

      // Shuffle available indexes
      const randomIndex = [...availableIndexes].sort(() => Math.random() - 0.5);
      const usedReferences = new Set<number>();

      // Modify shapes with references
      for (const index of randomIndex) {
        try {
          let referenceIndex: number;

          if (usedReferences.size < references.length) {
            const unusedReferences = Array.from(Array(references.length).keys()).filter((i) => !usedReferences.has(i));
            referenceIndex = unusedReferences[Math.floor(Math.random() * unusedReferences.length)];
            usedReferences.add(referenceIndex);
          } else {
            referenceIndex = Math.floor(Math.random() * references.length);
          }

          const { shape, text } = shapesWithText[index];
          const textRange = shape.textFrame.textRange;
          textRange.load("text");
          await context.sync();

          const trimmedText = text.trim();
          const newText = trimmedText.endsWith(".")
            ? trimmedText.slice(0, -1) + ` ${references[referenceIndex]}.`
            : trimmedText + ` ${references[referenceIndex]}.`;

          textRange.text = newText;
          await context.sync();
        } catch (error) {
          console.warn("Error updating shape:", error);
          continue;
        }
      }

      console.log("check 10");

      await context.sync();
      return `Successfully added ${usedReferences.size} references`;
    });
  } catch (error) {
    console.error("Error in analyzeDocument:", error);
    throw new Error(`Error modifying document: ${error.message}`);
  }
}

export async function removeReferences(): Promise<string> {
  try {
    return await PowerPoint.run(async (context) => {
      // Get all slides instead of selected slides
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      const citationPatterns = [
        /\((?:[^,()]+(,\s[^,()]+)*(?:,\sand\s[^,()]+)?)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+\sand\s[^,()]+)[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^()]+)\set\sal\.?[,\s]\s?\d{4}[a-z]?\)/g,
        /\((?:[^,()]+(,\s[^,()]+)*)[,\s]\s?\d{4}[a-z]?\)/g,
      ];

      let totalRemoved = 0;

      for (let slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        for (let shape of shapes.items) {
          if (shape.textFrame) {
            try {
              const textRange = shape.textFrame.textRange;
              textRange.load("text");
              await context.sync();

              let text = textRange.text;
              let hadMatch = false;

              for (const pattern of citationPatterns) {
                const matches = text.match(pattern) || [];
                if (matches.length > 0) {
                  text = text.replace(pattern, "").replace(/\s+\./g, ".");
                  hadMatch = true;
                  totalRemoved += matches.length;
                }
              }

              if (hadMatch) {
                // Instead of directly setting text, load the textRange again
                await context.sync();
                textRange.text = text;
                await context.sync();
                console.log("success had match");
              }
            } catch (error) {
              // Log error but continue processing other shapes
              console.warn("Error processing shape:", error);
              continue;
            }
          }
        }
      }

      // Final sync after all operations
      await context.sync();
      console.log("reach end");
      return `Removed ${totalRemoved} citations and cleaned up formatting`;
    });
  } catch (error) {
    console.error("Error in removeReferences:", error);
    throw new Error(`Error removing references: ${error.message}`);
  }
}
