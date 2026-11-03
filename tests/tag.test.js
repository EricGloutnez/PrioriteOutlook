/*
 * Tests de tag.js — exécuter avec : node tests/tag.test.js
 */
"use strict";
const assert = require("assert");
global.window = {};
require("../src/tag.js");
const T = window.PrioriteTag;

let ok = 0;
function test(nom, fn) { fn(); ok++; console.log("ok - " + nom); }

test("construire niveau seul", () => assert.strictEqual(T.construire("P1"), "[P1]"));
test("construire avec projet", () => assert.strictEqual(T.construire("P1", "M", "12345"), "[P1-M-12345]"));
test("refuser numéro non numérique", () => assert.throws(() => T.construire("P1", "M", "12a")));
test("refuser numéro sans type", () => assert.throws(() => T.construire("P2", null, "123")));
test("refuser niveau inconnu", () => assert.throws(() => T.construire("P4")));

test("lire balise simple", () => {
  const b = T.lire("[P2] Horaire");
  assert.strictEqual(b.niveau, "P2"); assert.strictEqual(b.numero, null);
});
test("lire balise projet après RE:", () => {
  const b = T.lire("RE: [URGENT-R-20871] Fuite");
  assert.deepStrictEqual([b.niveau, b.type, b.numero], ["URGENT", "R", "20871"]);
});
test("lire saisie manuelle en minuscules", () => assert.strictEqual(T.lire("[p3] test").niveau, "P3"));
test("ignorer texte sans balise", () => assert.strictEqual(T.lire("Projet P1 en cours"), null));

test("appliquer au début", () => assert.strictEqual(T.appliquer("Livraison", "[P1]"), "[P1] Livraison"));
test("appliquer sur objet vide", () => assert.strictEqual(T.appliquer("", "[P3]"), "[P3]"));
test("remplacer à sa position", () =>
  assert.strictEqual(T.appliquer("RE: [P2] Livraison", "[P1-M-12345]"), "RE: [P1-M-12345] Livraison"));
test("aucune double balise", () =>
  assert.strictEqual(T.appliquer("[P1-M-1] A", "[P1]"), "[P1] A"));

test("retirer balise", () => assert.strictEqual(T.retirer("TR: [P1-M-12345] Livraison"), "TR: Livraison"));
test("retirer sans balise", () => assert.strictEqual(T.retirer("Bonjour"), "Bonjour"));

console.log(`\n${ok} tests réussis`);
