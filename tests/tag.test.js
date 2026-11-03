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

// Construction
test("niveau seul", () => assert.strictEqual(T.construire("P1"), "[P1]"));
test("urgent seul", () => assert.strictEqual(T.construire("URGENT"), "[URGENT]"));
test("avec projet Multi", () => assert.strictEqual(T.construire("P1", "M", "11111"), "[M_11111_P1]"));
test("urgent avec projet Résidentiel", () => assert.strictEqual(T.construire("URGENT", "R", "11111"), "[R_11111_URGENT]"));
test("confidentiel seul", () => assert.strictEqual(T.construire("CONFIDENTIEL"), "[CONFIDENTIEL]"));
test("confidentiel en avant du projet", () => assert.strictEqual(T.construire("CONFIDENTIEL", "M", "11111"), "[CONFIDENTIEL_M_11111]"));
test("refuser numéro non numérique", () => assert.throws(() => T.construire("P1", "M", "12a")));
test("refuser numéro sans type", () => assert.throws(() => T.construire("P2", null, "123")));
test("refuser niveau inconnu", () => assert.throws(() => T.construire("P4")));

// Lecture
test("lire niveau seul", () => {
  const b = T.lire("[P2] Horaire");
  assert.deepStrictEqual([b.niveau, b.type, b.numero], ["P2", null, null]);
});
test("lire projet après RE:", () => {
  const b = T.lire("RE: [R_20871_URGENT] Fuite");
  assert.deepStrictEqual([b.niveau, b.type, b.numero], ["URGENT", "R", "20871"]);
});
test("lire confidentiel avec projet", () => {
  const b = T.lire("[CONFIDENTIEL_M_11111] Dossier");
  assert.deepStrictEqual([b.niveau, b.type, b.numero], ["CONFIDENTIEL", "M", "11111"]);
});
test("lire saisie manuelle en minuscules", () => assert.strictEqual(T.lire("[m_11111_p3] test").niveau, "P3"));
test("ignorer P1 sans crochets", () => assert.strictEqual(T.lire("P1-Fall Cottage Life Show"), null));
test("ignorer ancien format", () => assert.strictEqual(T.lire("[P1-R-17556] Message test"), null));

// Application
test("appliquer au début", () => assert.strictEqual(T.appliquer("Livraison", "[M_11111_P1]"), "[M_11111_P1] Livraison"));
test("appliquer sur objet vide", () => assert.strictEqual(T.appliquer("", "[P3]"), "[P3]"));
test("appliquer après RE:", () => assert.strictEqual(T.appliquer("RE: Livraison", "[P1]"), "RE: [P1] Livraison"));
test("appliquer après TR: RE:", () => assert.strictEqual(T.appliquer("TR: RE: Livraison", "[P2]"), "TR: RE: [P2] Livraison"));
test("remplacer à sa position", () =>
  assert.strictEqual(T.appliquer("RE: [P2] Livraison", "[M_11111_P1]"), "RE: [M_11111_P1] Livraison"));
test("aucune double balise", () =>
  assert.strictEqual(T.appliquer("[M_1_P1] A", "[CONFIDENTIEL]"), "[CONFIDENTIEL] A"));

// Retrait
test("retirer balise", () => assert.strictEqual(T.retirer("TR: [M_11111_P1] Livraison"), "TR: Livraison"));
test("retirer sans balise", () => assert.strictEqual(T.retirer("Bonjour"), "Bonjour"));

console.log(`\n${ok} tests réussis`);
