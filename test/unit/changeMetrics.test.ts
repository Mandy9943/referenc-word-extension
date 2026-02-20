import * as assert from "assert";
import "mocha";
import { calculateTextChangeMetrics } from "../../src/taskpane/changeMetrics";

/* global describe, it */

describe("changeMetrics", function () {
  it("counts complete lexical reuse as 0% changed", function () {
    const metrics = calculateTextChangeMetrics(
      ["Hello, world! Hello."],
      ["hello world hello"],
    );

    assert.strictEqual(metrics.totalOriginalWords, 3);
    assert.strictEqual(metrics.totalNewWords, 3);
    assert.strictEqual(metrics.reusedWords, 3);
    assert.strictEqual(metrics.wordsChanged, 0);
    assert.strictEqual(metrics.changePercent, 0);
    assert.strictEqual(metrics.reusePercent, 100);
    assert.strictEqual(metrics.addedWords, 0);
  });

  it("counts full replacement as 100% changed", function () {
    const metrics = calculateTextChangeMetrics(
      ["alpha beta gamma"],
      ["delta epsilon zeta"],
    );

    assert.strictEqual(metrics.totalOriginalWords, 3);
    assert.strictEqual(metrics.totalNewWords, 3);
    assert.strictEqual(metrics.reusedWords, 0);
    assert.strictEqual(metrics.wordsChanged, 3);
    assert.strictEqual(metrics.changePercent, 100);
    assert.strictEqual(metrics.reusePercent, 0);
    assert.strictEqual(metrics.addedWords, 3);
  });

  it("handles duplicates and partial overlap with frequency awareness", function () {
    const metrics = calculateTextChangeMetrics(
      ["cat cat dog"],
      ["cat dog bird bird"],
    );

    assert.strictEqual(metrics.totalOriginalWords, 3);
    assert.strictEqual(metrics.totalNewWords, 4);
    assert.strictEqual(metrics.reusedWords, 2);
    assert.strictEqual(metrics.wordsChanged, 1);
    assert.strictEqual(metrics.changePercent, 33);
    assert.strictEqual(metrics.reusePercent, 67);
    assert.strictEqual(metrics.addedWords, 2);
  });

  it("normalizes accented characters for fair matching", function () {
    const metrics = calculateTextChangeMetrics(
      ["Niño déjà résumé"],
      ["nino deja resume"],
    );

    assert.strictEqual(metrics.reusedWords, 3);
    assert.strictEqual(metrics.wordsChanged, 0);
    assert.strictEqual(metrics.changePercent, 0);
    assert.strictEqual(metrics.reusePercent, 100);
  });
});
