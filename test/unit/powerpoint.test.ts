import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { inferReferenceSlideIndexFromSlideTexts, insertText } from "../../src/taskpane/powerpoint";

/* global describe, global, it */

const shapes: any[] = [];

function createMockTextBox(text: string) {
  const shape = {
    text,
    fill: {
      setSolidColor: function () {
        /* no-op */
      },
    },
    lineFormat: {
      color: "",
      weight: 0,
      dashStyle: "",
    },
  };

  shapes.push(shape);
  return shape;
}

const selectedSlide = {
  shapes: {
    addTextBox: function (text: string) {
      return createMockTextBox(text);
    },
    items: shapes,
  },
};
const PowerPointMockData = {
  context: {
    presentation: {
      getSelectedSlides: function () {
        return {
          load: function () {
            /* no-op */
          },
          items: [selectedSlide],
        };
      },
    },
    sync: async function () {
      /* no-op */
    },
    slides: {
      items: [selectedSlide],
    },
  },
  onReady: async function () {},
  run: async function (callback) {
    await callback(this.context);
  },
};

describe(`PowerPoint`, function () {
  it("Inserts text", async function () {
    const officeMock = new OfficeMockObject(PowerPointMockData);
    (officeMock as any).ShapeLineDashStyle = { solid: "solid" };
    global.PowerPoint = officeMock as any;

    await insertText("Hello PowerPoint");

    assert.strictEqual(shapes[0].text, "Hello PowerPoint");
  });

  it("Infers reference slide without explicit references title", function () {
    const slideGroups = [
      [
        "Mental health in older adults is influenced by social isolation, chronic illness, and socioeconomic factors.",
        "This section introduces prevalence and risk factors.",
      ],
      [
        "Interventions include CBT, social prescribing, and integrated multidisciplinary care pathways.",
      ],
      [
        "1. Alzheimer's Association (2024). Facts and Figures. https://www.alz.org/facts",
        "2. WHO (2023). Mental health of older adults. Retrieved from https://www.who.int/news-room/fact-sheets",
        "3. Smith, J., & Doe, A. (2022). Journal of Geriatric Care, 10(2), 1-10. doi:10.1000/example",
      ],
    ];

    const inferredIndex = inferReferenceSlideIndexFromSlideTexts(slideGroups);
    assert.strictEqual(inferredIndex, 2);
  });

  it("Returns -1 when no reference-like slide exists", function () {
    const slideGroups = [
      [
        "Introduction to social care practice.",
        "Learning outcomes and module overview.",
      ],
      [
        "Case study findings show improvements in wellbeing and communication.",
      ],
    ];

    const inferredIndex = inferReferenceSlideIndexFromSlideTexts(slideGroups);
    assert.strictEqual(inferredIndex, -1);
  });
});
