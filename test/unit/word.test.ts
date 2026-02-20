import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { findReferenceStartIndexFromTexts, insertText } from "../../src/taskpane/word";

/* global describe, global, it, Word */

const WordMockData = {
  context: {
    document: {
      body: {
        paragraph: {
          text: "",
        },
        insertParagraph: function (paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  InsertLocation: {
    end: "End",
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

describe("Word", function () {
  it("Inserts text", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    global.Word = wordMock as any;

    await insertText("Hello Word");

    wordMock.context.document.body.paragraph.load("text");
    await wordMock.context.sync();

    assert.strictEqual(wordMock.context.document.body.paragraph.text, "Hello Word");
  });

  it("Finds numbered references header", function () {
    const index = findReferenceStartIndexFromTexts([
      "Introduction",
      "Body paragraph",
      "4. References",
      "Smith, J. (2021).",
    ]);

    assert.strictEqual(index, 2);
  });

  it("Uses last references header in document", function () {
    const index = findReferenceStartIndexFromTexts([
      "References",
      "Old list",
      "Appendix A",
      "Bibliography",
      "New list item",
    ]);

    assert.strictEqual(index, 3);
  });

  it("Does not match non-header mentions of references", function () {
    const index = findReferenceStartIndexFromTexts([
      "This section references prior studies and compares results.",
      "No bibliography heading here.",
    ]);

    assert.strictEqual(index, -1);
  });
});
