/*
 * Tests de tag.js — exécuter avec : node tests/tag.test.js
 */
"use strict";
const assert = require("assert");
global.window = {};
require("../src/tag.js");
const T = window.PrioriteTag;
const R = "\u{1F534}", O = "\u{1F7E0}", J = "\u{1F7E1}", V = "\u{1F7E2}", M = "\u{1F7E3}";

let ok = 0;
function test(nom, fn) { fn(); ok++; console.log("ok - " + nom); }

// Construction de la balise
test("niveau seul", () => assert.strictEqual(T.construire("P1"), "[P1]"));
test("avec projet Multi", () => assert.strictEqual(T.construire("P1", "M", "11111"), "[M_11111_P1]"));
test("urgent avec projet Résidentiel", () => assert.strictEqual(T.construire("URGENT", "R", "11111"), "[R_11111_URGENT]"));
test("confidentiel en avant du projet", () => assert.strictEqual(T.construire("CONFIDENTIEL", "M", "11111"), "[CONFIDENTIEL_M_11111]"));
test("refuser numéro non numérique", () => assert.throws(() => T.construire("P1", "M", "12a")));
test("refuser numéro sans type", () => assert.throws(() => T.construire("P2", null, "123")));
test("refuser niveau inconnu", () => assert.throws(() => T.construire("P4")));

// Marque complète (pastille + balise)
test("marque urgent", () => assert.strictEqual(T.marque("URGENT"), R + " [URGENT]"));
test("marque P1 projet", () => assert.strictEqual(T.marque("P1", "M", "12345"), O + " [M_12345_P1]"));
test("marque P2", () => assert.strictEqual(T.marque("P2"), J + " [P2]"));
test("marque P3", () => assert.strictEqual(T.marque("P3"), V + " [P3]"));
test("marque confidentiel", () => assert.strictEqual(T.marque("CONFIDENTIEL", "R", "1"), M + " [CONFIDENTIEL_R_1]"));

// Lecture
test("lire avec pastille", () => {
  const b = T.lire(O + " [M_12345_P1] Livraison");
  assert.deepStrictEqual([b.niveau, b.type, b.numero, b.debut, b.index], ["P1", "M", "12345", 0, 3]);
});
test("lire sans pastille (téléphone)", () => {
  const b = T.lire("[P2] Horaire");
  assert.deepStrictEqual([b.niveau, b.debut, b.index], ["P2", 0, 0]);
});
test("lire projet après RE:", () => {
  const b = T.lire("RE: " + R + " [R_20871_URGENT] Fuite");
  assert.deepStrictEqual([b.niveau, b.type, b.numero], ["URGENT", "R", "20871"]);
});
test("ignorer P1 sans crochets", () => assert.strictEqual(T.lire("P1-Fall Cottage Life Show"), null));

// Application
test("appliquer au début", () => assert.strictEqual(T.appliquer("Livraison", T.marque("P1", "M", "11111")), O + " [M_11111_P1] Livraison"));
test("appliquer sur objet vide", () => assert.strictEqual(T.appliquer("", T.marque("P3")), V + " [P3]"));
test("appliquer après RE:", () => assert.strictEqual(T.appliquer("RE: Livraison", T.marque("P1")), "RE: " + O + " [P1] Livraison"));
test("remplacer pastille et balise", () =>
  assert.strictEqual(T.appliquer("RE: " + J + " [P2] Livraison", T.marque("URGENT")), "RE: " + R + " [URGENT] Livraison"));
test("ajouter pastille à une balise tapée à la main", () =>
  assert.strictEqual(T.appliquer("[P2] Horaire", T.marque("P2")), J + " [P2] Horaire"));
test("aucune double pastille", () =>
  assert.strictEqual(T.appliquer(O + " Livraison", T.marque("P1")), O + " [P1] Livraison"));

// Retrait
test("retirer pastille et balise", () => assert.strictEqual(T.retirer("TR: " + O + " [M_11111_P1] Livraison"), "TR: Livraison"));
test("retirer balise sans pastille", () => assert.strictEqual(T.retirer("[P1] Livraison"), "Livraison"));
test("retirer sans balise", () => assert.strictEqual(T.retirer("Bonjour"), "Bonjour"));

console.log(`\n${ok} tests réussis`);
