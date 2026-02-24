import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { findReferenceStartIndexFromTexts, insertText, sanitizeParaphraseOutputText } from "../../src/taskpane/word";

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

  it("Infers headerless reference block from tail reference-like entries", function () {
    const index = findReferenceStartIndexFromTexts([
      "Introduction paragraph with context and motivation for the topic.",
      "Body analysis continues with findings and interpretation.",
      "Arab News (2025) Saudi Arabia market update. Available at: https://example.com/arab-news",
      "CNBC (2025) Industry outlook and growth indicators. Available at: https://example.com/cnbc-outlook",
    ]);

    assert.strictEqual(index, 2);
  });

  it("Does not infer references from normal narrative tail text", function () {
    const index = findReferenceStartIndexFromTexts([
      "Tourism trends are evolving rapidly across digital channels in 2025.",
      "This paragraph explains behavior shifts and campaign tactics in practical terms.",
      "Results indicate that visual storytelling improves conversion in multiple segments.",
      "Recommendations focus on execution quality and continuous optimization.",
    ]);

    assert.strictEqual(index, -1);
  });

  it("Strips internal delimiter token from paraphrase output", function () {
    const cleaned = sanitizeParaphraseOutputText("A critical evaluation of qbpdelim123 SDL models.");
    assert.strictEqual(cleaned.includes("qbpdelim123"), false);
    assert.strictEqual(cleaned, "A critical evaluation of SDL models.");
  });

  it("Strips internal delimiter token case-insensitively", function () {
    const cleaned = sanitizeParaphraseOutputText("alpha QBPDELIM123 beta");
    assert.strictEqual(cleaned, "alpha beta");
  });
});
