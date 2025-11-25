import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { insertText } from "../../src/taskpane/powerpoint";

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
});
