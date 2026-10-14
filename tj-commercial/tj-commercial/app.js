/* Terre & Jardins, tableau de bord commercial : page data-tools.
   Filtres, vues en onglets, fiches article et client, export Excel.
   Aucune dépendance externe, tout passe par dt.js (dtInit, dtFetch, dtDownload). */
(function () {
    'use strict';

    var CONFIG = {
        "outil": "tj-commercial",
        "titre": "Terre & Jardins, tableau de bord commercial",
        "fichier": "TJ_commercial.xlsx",
        "periodes": [
            {
                "k": "hier",
                "l": "Hier"
            },
            {
                "k": "7j",
                "l": "7 derniers jours"
            },
            {
                "k": "mois",
                "l": "Mois en cours"
            },
            {
                "k": "moisp",
                "l": "Mois dernier"
            },
            {
                "k": "trim",
                "l": "Trimestre"
            },
            {
                "k": "juin",
                "l": "Depuis le 1er juin"
            }
        ],
        "onglets": [
            {
                "k": "vue",
                "l": "Vue d'ensemble"
            },
            {
                "k": "articles",
                "l": "Articles"
            },
            {
                "k": "clients",
                "l": "Clients Pro"
            },
            {
                "k": "temps",
                "l": "Temps"
            },
            {
                "k": "plateformes",
                "l": "Plateformes"
            },
            {
                "k": "comptoir",
                "l": "Comptoir"
            },
            {
                "k": "controle",
                "l": "Contrôle"
            }
        ],
        "plateformes": true,
        "segments": false,
        "objectifs": true
    };

    /* =====================================================================
       RENDU COMMUN DES TABLEAUX DE BORD (tj-commercial, ecobois-commercial)
       Blocs renvoyés par le moteur : note, kpis, table, chart, hbar, heat, objectifs.
       Aucune dépendance externe, aucun style inline (el.style en JS uniquement).
       ===================================================================== */

    var SVG_NS = 'http://www.w3.org/2000/svg';
    var TAILLES_PAGE = [10, 25, 50, 100, 0];      // 0 = tout

    // ---------- Petits outils ----------
    function $(id) { return document.getElementById(id); }

    function el(tag, classe, texte) {
        var e = document.createElement(tag);
        if (classe) e.className = classe;
        if (texte !== undefined && texte !== null) e.textContent = texte;
        return e;
    }

    function svgEl(tag, attributs) {
        var e = document.createElementNS(SVG_NS, tag);
        for (var cle in attributs) e.setAttribute(cle, String(attributs[cle]));
        return e;
    }

    function afficher(noeud, visible) {
        if (noeud) noeud.classList.toggle('dt-cache', !visible);
    }

    function normaliser(t) {
        return String(t === null || t === undefined ? '' : t)
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    }

    function nf(x, a, b) {
        return Number(x).toLocaleString('fr-FR', { minimumFractionDigits: a, maximumFractionDigits: b });
    }

    function estNombre(f) {
        return ['eur', 'eur0', 'int', 'num', 'pct', 'days', 'sem'].indexOf(f) !== -1;
    }

    function fmt(v, f, grand, signe) {
        if (v === null || v === undefined || v === '' || (typeof v === 'number' && !isFinite(v))) return grand ? 'n.d.' : '';
        var sg = function (x) { return signe && x > 0 ? '+' : ''; };
        if (f === 'eur') return sg(v) + nf(v, 2, 2) + ' €';
        if (f === 'eur0') return sg(v) + nf(Math.round(v), 0, 0) + ' €';
        if (f === 'int') return nf(v, 0, 0);
        if (f === 'num') return nf(v, 0, 2);
        if (f === 'pct') return sg(v) + nf(v * 100, 1, 1) + ' %';
        if (f === 'days') return nf(v, 0, 1) + ' j';
        if (f === 'sem') return nf(v, 0, 1) + ' sem.';
        if (f === 'date') {
            var s = String(v);
            return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : s;
        }
        return String(v);
    }

    function dateHeure(iso) {
        var d = new Date(iso);
        if (isNaN(d.getTime())) return '';
        return d.toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    }

    function sousTitre(s) {
        return String(s.t).replace('{v}', s.v === undefined ? '' : fmt(s.v, s.f, true));
    }

    function ajouterJours(s, n) {
        var t = Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) + n * 864e5;
        return new Date(t).toISOString().slice(0, 10);
    }

    // Écart d'un indicateur avec sa référence (p = période précédente, n1 = même période N-1)
    function variation(it, cle) {
        var ref = it[cle];
        var lib = cle === 'n1' ? 'N-1' : 'préc.';
        if (ref === null || ref === undefined || it.v === null || it.v === undefined || typeof it.v !== 'number') return null;
        var d, t;
        if (it.f === 'pct') {
            d = (it.v - ref) * 100;
            t = (d >= 0 ? '+' : '') + nf(d, 1, 1) + ' pt';
        } else {
            if (!ref) return { classe: 'tb-neutre', texte: lib + ' 0' };
            d = (it.v - ref) / Math.abs(ref);
            t = (d >= 0 ? '+' : '') + nf(d * 100, 1, 1) + ' %';
        }
        var fleche = Math.abs(d) < 1e-9 ? '■' : d > 0 ? '▲' : '▼';
        var bon = it.good === undefined ? 1 : it.good;
        var classe = !bon || Math.abs(d) < 1e-9 ? 'tb-neutre' : (d > 0) === (bon > 0) ? 'dt-hausse' : 'dt-baisse';
        return { classe: classe, texte: fleche + ' ' + t + ' vs ' + lib + ' (' + fmt(ref, it.f, true) + ')' };
    }

    // ---------- Sélecteur multiple avec recherche sans accents ----------
    // options : [{ code, titre }] ; aucune valeur choisie = « Tous »
    function creerSelecteur(hote, reglages) {
        var MAX_AFFICHEES = 300;
        var etatSel = { options: [], valeurs: [], ouvert: false };

        var racine = el('div', 'tb-sel');
        var libelle = el('div', 'tb-sel-lib', reglages.libelle);
        var champ = el('button', 'tb-sel-champ');
        champ.type = 'button';
        champ.setAttribute('aria-haspopup', 'listbox');
        champ.setAttribute('aria-expanded', 'false');
        var puces = el('span', 'tb-sel-puces');
        champ.append(puces, el('span', 'tb-sel-fleche', '▾'));

        var panneau = el('div', 'tb-sel-panneau dt-cache');
        var recherche = el('input', 'dt-input tb-sel-recherche');
        recherche.type = 'search';
        recherche.placeholder = 'Rechercher…';
        recherche.setAttribute('aria-label', 'Rechercher ' + reglages.libelle);
        var barre = el('div', 'tb-sel-actions');
        var btnTout = el('button', 'dt-btn dt-btn-sec', 'Tout');
        btnTout.type = 'button';
        var btnFermer = el('button', 'dt-btn dt-btn-sec', 'Fermer');
        btnFermer.type = 'button';
        barre.append(btnTout, el('span', 'dt-espace'), btnFermer);
        var liste = el('div', 'tb-sel-liste');
        liste.setAttribute('role', 'listbox');
        liste.setAttribute('aria-multiselectable', 'true');
        var note = el('div', 'tb-sel-note dt-muet');
        panneau.append(recherche, barre, liste, note);
        racine.append(libelle, champ, panneau);
        hote.textContent = '';
        hote.append(racine);

        function titreDe(code) {
            for (var i = 0; i < etatSel.options.length; i++) {
                if (etatSel.options[i].code === code) return etatSel.options[i].titre;
            }
            return code;
        }

        function dessinerPuces() {
            puces.textContent = '';
            if (etatSel.valeurs.length === 0) {
                puces.append(el('span', 'dt-muet', 'Tous'));
                return;
            }
            etatSel.valeurs.forEach(function (code) {
                var puce = el('span', 'tb-puce', titreDe(code) + ' ×');
                puce.title = 'Retirer ' + titreDe(code);
                puce.addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    etatSel.valeurs = etatSel.valeurs.filter(function (v) { return v !== code; });
                    dessinerPuces();
                    dessinerListe();
                    reglages.onChange();
                });
                puces.append(puce);
            });
        }

        function dessinerListe() {
            var mots = normaliser(recherche.value).trim().split(/\s+/).filter(Boolean);
            var trouvees = etatSel.options.filter(function (o) {
                var texte = normaliser(o.titre);
                return mots.every(function (m) { return texte.indexOf(m) !== -1; });
            });
            liste.textContent = '';
            trouvees.slice(0, MAX_AFFICHEES).forEach(function (o) {
                var choisi = etatSel.valeurs.indexOf(o.code) !== -1;
                var ligne = el('div', 'tb-sel-option' + (choisi ? ' tb-sel-choisie' : ''));
                ligne.setAttribute('role', 'option');
                ligne.setAttribute('aria-selected', choisi ? 'true' : 'false');
                ligne.append(el('span', 'tb-sel-case', choisi ? '☑' : '☐'), el('span', 'tb-sel-texte', o.titre));
                ligne.addEventListener('click', function () {
                    if (etatSel.valeurs.indexOf(o.code) === -1) etatSel.valeurs.push(o.code);
                    else etatSel.valeurs = etatSel.valeurs.filter(function (v) { return v !== o.code; });
                    dessinerPuces();
                    dessinerListe();
                    reglages.onChange();
                });
                liste.append(ligne);
            });
            if (trouvees.length === 0) note.textContent = 'Aucun résultat';
            else if (trouvees.length > MAX_AFFICHEES) note.textContent = MAX_AFFICHEES + ' affichés sur ' + nf(trouvees.length, 0, 0) + ', affine la recherche';
            else note.textContent = nf(trouvees.length, 0, 0) + ' choix' + (etatSel.valeurs.length ? ', ' + etatSel.valeurs.length + ' sélectionné(s)' : '');
        }

        function ouvrir(oui) {
            etatSel.ouvert = oui;
            afficher(panneau, oui);
            champ.setAttribute('aria-expanded', oui ? 'true' : 'false');
            if (oui) {
                dessinerListe();
                recherche.focus();
            }
        }

        champ.addEventListener('click', function () { ouvrir(!etatSel.ouvert); });
        recherche.addEventListener('input', dessinerListe);
        recherche.addEventListener('keydown', function (ev) {
            if (ev.key === 'Escape') { ouvrir(false); champ.focus(); }
        });
        btnFermer.addEventListener('click', function () { ouvrir(false); });
        btnTout.addEventListener('click', function () {
            etatSel.valeurs = [];
            recherche.value = '';
            dessinerPuces();
            dessinerListe();
            reglages.onChange();
        });
        document.addEventListener('click', function (ev) {
            if (etatSel.ouvert && !racine.contains(ev.target)) ouvrir(false);
        });

        dessinerPuces();

        return {
            definirOptions: function (options) {
                // les valeurs choisies restent proposées même si la période ne les contient plus
                var connues = {};
                (options || []).forEach(function (o) { connues[o.code] = true; });
                var anciennes = etatSel.options;
                etatSel.options = (options || []).slice();
                etatSel.valeurs.forEach(function (v) {
                    if (connues[v]) return;
                    var ancienne = anciennes.filter(function (o) { return o.code === v; })[0];
                    etatSel.options.unshift(ancienne || { code: v, titre: v });
                });
                dessinerPuces();
                if (etatSel.ouvert) dessinerListe();
            },
            valeurs: function () { return etatSel.valeurs.slice(); },
            vider: function () {
                etatSel.valeurs = [];
                dessinerPuces();
                if (etatSel.ouvert) dessinerListe();
            }
        };
    }

    // ---------- Graphique en barres et courbes ----------
    function libelleCourt(v, f) {
        if (f === 'pct') return Math.round(v * 100) + ' %';
        var a = Math.abs(v), e = f === 'eur' || f === 'eur0' ? ' €' : '';
        if (a >= 1e6) return nf(v / 1e6, 0, 1) + ' M' + e;
        if (a >= 1e4) return nf(v / 1e3, 0, 0) + ' k' + e;
        if (a >= 1e3) return nf(v / 1e3, 0, 1) + ' k' + e;
        return nf(v, 0, 1) + e;
    }

    function echelle(bas, haut, n) {
        if (haut - bas < 1e-9) haut = bas + 1;
        var brut = (haut - bas) / n, mag = Math.pow(10, Math.floor(Math.log10(brut))), r = brut / mag;
        var pas = (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * mag;
        return { bas: Math.floor(bas / pas + 1e-9) * pas, haut: Math.ceil(haut / pas - 1e-9) * pas, pas: pas };
    }

    function dessinerGraphique(b) {
        var W = (b.w || 12) >= 12 ? 1000 : 540, H = 270;
        var n = (b.cats || []).length;
        if (!n || !b.series || !b.series.length) return el('p', 'dt-vide', 'Aucune donnée');

        var barres = b.series.filter(function (s) { return s.kind !== 'line'; });
        var lignes1 = b.series.filter(function (s) { return s.kind === 'line' && s.ax !== 2; });
        var lignes2 = b.series.filter(function (s) { return s.ax === 2; });
        var Lm = 62, Rm = lignes2.length ? 58 : 14, Tm = 10, Bm = 44, cw = W - Lm - Rm, ch = H - Tm - Bm;

        var mx = 0, mn = 0;
        for (var i = 0; i < n; i++) {
            if (b.stacked) {
                var p = 0, q = 0;
                barres.forEach(function (s) { var v = +s.d[i] || 0; if (v > 0) p += v; else q += v; });
                mx = Math.max(mx, p); mn = Math.min(mn, q);
            } else {
                barres.forEach(function (s) { var v = +s.d[i] || 0; mx = Math.max(mx, v); mn = Math.min(mn, v); });
            }
            lignes1.forEach(function (s) { var v = s.d[i]; if (v !== null && v !== undefined) { mx = Math.max(mx, v); mn = Math.min(mn, v); } });
        }
        var A = echelle(mn, mx, 5);
        var y = function (v) { return Tm + ch - ((v - A.bas) / (A.haut - A.bas)) * ch; };
        var A2 = null, y2 = null;
        if (lignes2.length) {
            var m2 = 0;
            lignes2.forEach(function (s) { s.d.forEach(function (v) { if (v > m2) m2 = v; }); });
            A2 = b.f2 === 'pct' ? { bas: 0, haut: Math.max(1, Math.ceil(m2 * 4) / 4), pas: 0.25 } : echelle(0, m2 || 1, 5);
            y2 = function (v) { return Tm + ch - ((v - A2.bas) / (A2.haut - A2.bas)) * ch; };
        }
        var bw = cw / n;
        var x = function (k) { return Lm + bw * k + bw / 2; };

        var svg = svgEl('svg', { viewBox: '0 0 ' + W + ' ' + H, width: '100%', preserveAspectRatio: 'xMidYMid meet', 'font-size': 11, 'font-family': 'Arial, sans-serif', role: 'img', 'aria-label': b.title || 'Graphique' });
        svg.classList.add('tb-svg');

        for (var v = A.bas; v <= A.haut + A.pas / 2; v += A.pas) {
            var yy = y(v).toFixed(1);
            svg.append(svgEl('line', { x1: Lm, x2: W - Rm, y1: yy, y2: yy, stroke: Math.abs(v) < 1e-9 ? '#9A8C6E' : '#D7DCD3', 'stroke-width': 1 }));
            var t = svgEl('text', { x: Lm - 6, y: +yy + 4, 'text-anchor': 'end', fill: '#5C645D' });
            t.textContent = libelleCourt(v, b.f);
            svg.append(t);
        }
        if (A2) {
            for (var v2 = A2.bas; v2 <= A2.haut + A2.pas / 2; v2 += A2.pas) {
                var t2 = svgEl('text', { x: W - Rm + 6, y: (y2(v2) + 4).toFixed(1), fill: '#6B5E43' });
                t2.textContent = libelleCourt(v2, b.f2);
                svg.append(t2);
            }
        }

        var ecart = Math.min(bw * 0.22, 8), interieur = bw - ecart;
        var formatBulle = b.f === 'eur0' ? 'eur' : b.f;
        function bulle(parent, texte) {
            var titre = svgEl('title', {});
            titre.textContent = texte;
            parent.append(titre);
        }
        for (var k = 0; k < n; k++) {
            var cat = String(b.cats[k]);
            if (b.stacked) {
                var haut = 0, bas = 0;
                barres.forEach(function (s) {
                    var val = +s.d[k] || 0;
                    if (!val) return;
                    var a = val > 0 ? haut : bas, z = a + val;
                    if (val > 0) haut = z; else bas = z;
                    var r = svgEl('rect', { x: (Lm + bw * k + ecart / 2).toFixed(1), y: Math.min(y(a), y(z)).toFixed(1), width: Math.max(1, interieur).toFixed(1), height: Math.max(0.5, Math.abs(y(a) - y(z))).toFixed(1), fill: s.c || '#2F5D3A' });
                    bulle(r, cat + ' · ' + s.n + ' : ' + fmt(val, formatBulle));
                    svg.append(r);
                });
            } else {
                var sw = interieur / Math.max(1, barres.length);
                barres.forEach(function (s, rang) {
                    var val = +s.d[k] || 0;
                    if (!val) return;
                    var r = svgEl('rect', { x: (Lm + bw * k + ecart / 2 + sw * rang).toFixed(1), y: Math.min(y(0), y(val)).toFixed(1), width: Math.max(1, sw - 1).toFixed(1), height: Math.max(0.5, Math.abs(y(val) - y(0))).toFixed(1), fill: s.c || '#2F5D3A' });
                    bulle(r, cat + ' · ' + s.n + ' : ' + fmt(val, formatBulle));
                    svg.append(r);
                });
            }
        }

        [[lignes1, y, b.f], [lignes2, y2, b.f2]].forEach(function (groupe) {
            var liste = groupe[0], fy = groupe[1], ff = groupe[2] === 'eur0' ? 'eur' : groupe[2];
            liste.forEach(function (s) {
                var d = '', stylo = false;
                for (var i2 = 0; i2 < n; i2++) {
                    var val = s.d[i2];
                    if (val === null || val === undefined) { stylo = false; continue; }
                    d += (stylo ? 'L' : 'M') + x(i2).toFixed(1) + ' ' + fy(val).toFixed(1) + ' ';
                    stylo = true;
                }
                var attrs = { d: d, fill: 'none', stroke: s.c || '#9A8C6E', 'stroke-width': 2 };
                if (s.da) attrs['stroke-dasharray'] = s.da;
                svg.append(svgEl('path', attrs));
                if (n <= 62) {
                    for (var i3 = 0; i3 < n; i3++) {
                        var val3 = s.d[i3];
                        if (val3 === null || val3 === undefined) continue;
                        var c = svgEl('circle', { cx: x(i3).toFixed(1), cy: fy(val3).toFixed(1), r: 2.5, fill: s.c || '#9A8C6E' });
                        bulle(c, String(b.cats[i3]) + ' · ' + s.n + ' : ' + fmt(val3, ff));
                        svg.append(c);
                    }
                }
            });
        });

        var chaque = Math.max(1, Math.ceil(n / Math.floor(cw / 46)));
        for (var j = 0; j < n; j += chaque) {
            var lib = String(b.cats[j]);
            var tx = svgEl('text', { x: x(j).toFixed(1), y: H - Bm + 16, 'text-anchor': 'middle', fill: '#5C645D' });
            tx.textContent = lib.length > 9 ? lib.slice(0, 8) + '…' : lib;
            svg.append(tx);
        }
        if (b.catT) {
            var tt = svgEl('text', { x: Lm + cw / 2, y: H - 6, 'text-anchor': 'middle', fill: '#9A8C6E' });
            tt.textContent = b.catT;
            svg.append(tt);
        }
        return svg;
    }

    // ---------- Blocs ----------
    // contexte : { tris: {}, pages: {}, tailles: {}, cleBase, ouvrirFiche(kind, code), saisirObjectifs(bloc), admin }
    function largeur(b, defaut) {
        var w = b.w || defaut || 12;
        return 'tb-w' + (w >= 12 ? 12 : w >= 6 ? 6 : 4);
    }

    function blocNote(b) {
        var niveau = b.lvl === 'warn' ? 'dt-msg-alerte' : b.lvl === 'err' ? 'dt-msg-erreur' : 'dt-msg-info';
        return el('div', 'dt-msg ' + niveau + ' tb-w12', b.text);
    }

    function blocKpis(b) {
        var carte = el('section', 'dt-carte tb-bloc ' + largeur(b, 12));
        if (b.title) carte.append(el('h3', 'tb-titre', b.title));
        var zone = el('div', 'dt-kpis tb-kpis');
        b.items.forEach(function (it) {
            var k = el('div', 'dt-kpi');
            k.append(el('div', 'dt-kpi-lib', it.l), el('div', 'dt-kpi-val', fmt(it.v, it.f, true)));
            ['p', 'n1'].forEach(function (cle) {
                var d = variation(it, cle);
                if (d) k.append(el('div', 'tb-var ' + d.classe, d.texte));
            });
            if (it.sub) k.append(el('div', 'dt-kpi-sub', sousTitre(it.sub)));
            zone.append(k);
        });
        carte.append(zone);
        return carte;
    }

    function blocTable(b, ctx) {
        var cle = ctx.cleBase + b.id;
        if (!(cle in ctx.tris)) ctx.tris[cle] = b.sort ? { cle: b.sort, sens: -1 } : null;
        if (!(cle in ctx.tailles)) ctx.tailles[cle] = b.nopage ? 0 : (b.pp || 10);
        if (!(cle in ctx.pages)) ctx.pages[cle] = 1;

        var carte = el('section', 'dt-carte tb-bloc ' + largeur(b, 12) + (b.warn ? ' tb-alerte' : ''));
        var entete = el('div', 'dt-ligne tb-entete-bloc');
        var titre = el('h3', 'tb-titre', b.title);
        var compte = el('span', 'dt-muet', nf(b.rows.length, 0, 0) + ' lignes' + (b.tronque ? ' affichées sur ' + nf(b.total, 0, 0) + ' (toutes dans l’export Excel)' : ''));
        entete.append(titre, el('span', 'dt-espace'), compte);
        carte.append(entete);

        var zone = el('div', 'dt-table-zone');
        var table = el('table', 'dt-table');
        var thead = el('thead');
        var trh = el('tr');
        var tbody = el('tbody');
        table.append(thead, tbody);
        thead.append(trh);
        zone.append(table);
        carte.append(zone);

        var pied = el('div', 'dt-ligne tb-pagination');
        var selTaille = el('select', 'dt-input tb-taille');
        selTaille.setAttribute('aria-label', 'Lignes par page');
        TAILLES_PAGE.forEach(function (t) {
            var o = el('option', null, t === 0 ? 'Tout' : String(t));
            o.value = String(t);
            selTaille.append(o);
        });
        var btnPrec = el('button', 'dt-btn dt-btn-sec', 'Précédente');
        btnPrec.type = 'button';
        var infoPage = el('span', 'dt-muet');
        var btnSuiv = el('button', 'dt-btn dt-btn-sec', 'Suivante');
        btnSuiv.type = 'button';
        pied.append(el('span', 'dt-muet', 'Lignes par page'), selTaille, el('span', 'dt-espace'), btnPrec, infoPage, btnSuiv);
        carte.append(pied);

        function lignesTriees() {
            var lignes = b.rows.slice();
            var tri = ctx.tris[cle];
            if (!tri) return lignes;
            lignes.sort(function (a, c) {
                var va = a[tri.cle], vc = c[tri.cle];
                var videA = va === null || va === undefined || va === '';
                var videC = vc === null || vc === undefined || vc === '';
                if (videA && videC) return 0;
                if (videA) return 1;
                if (videC) return -1;
                if (typeof va === 'number' && typeof vc === 'number') return (va - vc) * tri.sens;
                return String(va).localeCompare(String(vc), 'fr', { numeric: true }) * tri.sens;
            });
            return lignes;
        }

        function dessinerEntetes() {
            trh.textContent = '';
            var tri = ctx.tris[cle];
            b.cols.forEach(function (col) {
                var th = el('th', estNombre(col.f) ? 'dt-num' : null);
                var actif = tri && tri.cle === col.k;
                th.setAttribute('aria-sort', actif ? (tri.sens > 0 ? 'ascending' : 'descending') : 'none');
                var bouton = el('button', 'tb-th', col.t + (actif ? (tri.sens > 0 ? ' ▲' : ' ▼') : ''));
                bouton.type = 'button';
                bouton.addEventListener('click', function () {
                    var t = ctx.tris[cle];
                    if (t && t.cle === col.k) t.sens = -t.sens;
                    else ctx.tris[cle] = { cle: col.k, sens: estNombre(col.f) || col.f === 'date' ? -1 : 1 };
                    ctx.pages[cle] = 1;
                    dessinerEntetes();
                    dessinerCorps();
                });
                th.append(bouton);
                trh.append(th);
            });
        }

        function dessinerCorps() {
            var lignes = lignesTriees();
            var taille = ctx.tailles[cle];
            var pages = taille ? Math.max(1, Math.ceil(lignes.length / taille)) : 1;
            if (ctx.pages[cle] > pages) ctx.pages[cle] = pages;
            var debut = taille ? (ctx.pages[cle] - 1) * taille : 0;
            var visibles = taille ? lignes.slice(debut, debut + taille) : lignes;

            tbody.textContent = '';
            if (visibles.length === 0) {
                var tr0 = el('tr');
                var td0 = el('td', 'dt-vide', 'Aucune donnée');
                td0.colSpan = b.cols.length;
                tr0.append(td0);
                tbody.append(tr0);
            }
            visibles.forEach(function (ligne) {
                var tr = el('tr', b.click ? 'dt-cliquable' : null);
                if (b.click) {
                    tr.tabIndex = 0;
                    var ouvrir = function () {
                        var code = ligne[b.clickKey || 'code'];
                        if (code) ctx.ouvrirFiche(b.click, code);
                    };
                    tr.addEventListener('click', ouvrir);
                    tr.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') ouvrir(); });
                }
                b.cols.forEach(function (col) {
                    var v = ligne[col.k];
                    var classes = [];
                    if (estNombre(col.f)) classes.push('dt-num');
                    if (b.signed === col.k && typeof v === 'number') {
                        if (v > 0.0005) classes.push('tb-pos');
                        else if (v < -0.0005) classes.push('tb-neg');
                    }
                    if (col.f === 'code') classes.push('tb-code');
                    tr.append(el('td', classes.join(' ') || null, fmt(v, col.f, false, b.signed === col.k)));
                });
                tbody.append(tr);
            });

            afficher(pied, !b.nopage && b.rows.length > 10);
            selTaille.value = String(taille);
            infoPage.textContent = ctx.pages[cle] + ' / ' + pages;
            btnPrec.disabled = ctx.pages[cle] <= 1;
            btnSuiv.disabled = ctx.pages[cle] >= pages;
        }

        selTaille.addEventListener('change', function () {
            ctx.tailles[cle] = Number(selTaille.value);
            ctx.pages[cle] = 1;
            dessinerCorps();
        });
        btnPrec.addEventListener('click', function () { ctx.pages[cle] -= 1; dessinerCorps(); });
        btnSuiv.addEventListener('click', function () { ctx.pages[cle] += 1; dessinerCorps(); });

        dessinerEntetes();
        dessinerCorps();
        return carte;
    }

    function blocGraphique(b) {
        var carte = el('section', 'dt-carte tb-bloc ' + largeur(b, 12));
        carte.append(el('h3', 'tb-titre', b.title));
        var legende = el('div', 'tb-legende');
        (b.series || []).forEach(function (s) {
            var cle = el('i', s.kind === 'line' ? 'tb-cle tb-cle-ligne' : 'tb-cle');
            cle.style.background = s.c || '#2F5D3A';
            var item = el('span', 'tb-legende-item');
            item.append(cle, document.createTextNode(s.n));
            legende.append(item);
        });
        var zone = el('div', 'dt-graph');
        zone.append(dessinerGraphique(b));
        carte.append(legende, zone);
        return carte;
    }

    function blocBarres(b) {
        var carte = el('section', 'dt-carte tb-bloc ' + largeur(b, 6));
        carte.append(el('h3', 'tb-titre', b.title));
        if (!b.rows.length) {
            carte.append(el('p', 'dt-vide', 'Aucune donnée'));
            return carte;
        }
        var max = 0;
        b.rows.forEach(function (r) { max = Math.max(max, r.v || 0); });
        b.rows.forEach(function (r) {
            var ligne = el('div', 'tb-hb');
            var lib = el('div', 'tb-hb-lib', r.l);
            lib.title = r.l;
            var piste = el('div', 'tb-hb-piste');
            var barre = el('div', 'tb-hb-barre');
            barre.style.width = (max ? Math.max(0, r.v) / max * 100 : 0) + '%';
            piste.append(barre);
            var valeur = el('div', 'tb-hb-val', fmt(r.v, b.f));
            valeur.append(el('small', null, fmt(r.p, 'pct')));
            ligne.append(lib, piste, valeur);
            carte.append(ligne);
        });
        return carte;
    }

    function blocChaleur(b) {
        var carte = el('section', 'dt-carte tb-bloc ' + largeur(b, 12));
        var entete = el('div', 'dt-ligne');
        entete.append(el('h3', 'tb-titre', b.title), el('span', 'dt-espace'), el('span', 'dt-muet', 'couleur : CA HT, chiffre : tickets'));
        carte.append(entete);
        var max = 0;
        b.m.forEach(function (ligne) { ligne.forEach(function (v) { if (v > max) max = v; }); });
        var zone = el('div', 'dt-table-zone');
        var table = el('table', 'tb-chaleur');
        var trh = el('tr');
        trh.append(el('th'));
        b.cols.forEach(function (c) { trh.append(el('th', null, c)); });
        var thead = el('thead');
        thead.append(trh);
        var tbody = el('tbody');
        b.rows.forEach(function (jour, i) {
            var tr = el('tr');
            tr.append(el('th', null, jour));
            b.cols.forEach(function (c, k) {
                var v = b.m[i][k], nb = b.n[i][k];
                var td = el('td', null, nb ? String(nb) : '');
                td.title = jour + ' ' + c + ' : ' + fmt(v, 'eur') + ', ' + (nb || 0) + ' tickets';
                if (v && max) {
                    var a = Math.min(1, v / max);
                    td.style.background = 'rgba(47,93,58,' + (0.08 + a * 0.85).toFixed(3) + ')';
                    td.style.color = a > 0.55 ? '#fff' : '#1E3A25';
                }
                tr.append(td);
            });
            tbody.append(tr);
        });
        table.append(thead, tbody);
        zone.append(table);
        carte.append(zone);
        return carte;
    }

    function blocObjectifs(b, ctx) {
        var carte = el('section', 'dt-carte tb-bloc tb-w12');
        var entete = el('div', 'dt-ligne');
        entete.append(el('h3', 'tb-titre', b.title + ' : ' + b.moisLib), el('span', 'dt-espace'));
        if (ctx.admin && ctx.saisirObjectifs) {
            var bouton = el('button', 'dt-btn dt-btn-sec', 'Saisir les objectifs');
            bouton.type = 'button';
            bouton.addEventListener('click', function () { ctx.saisirObjectifs(b); });
            entete.append(bouton);
        }
        carte.append(entete);
        var zone = el('div', 'dt-table-zone');
        var table = el('table', 'dt-table');
        var trh = el('tr');
        ['Plateforme', 'Réalisé à date', 'Projection fin de mois', 'Objectif', 'Atteinte projetée', 'Atteinte à date'].forEach(function (t, i) {
            trh.append(el('th', i > 0 && i !== 4 ? 'dt-num' : null, t));
        });
        var thead = el('thead');
        thead.append(trh);
        var tbody = el('tbody');
        b.rows.forEach(function (r) {
            var tr = el('tr', r.total ? 'tb-total' : null);
            tr.append(el('td', null, r.pf), el('td', 'dt-num', fmt(r.real, 'eur0')), el('td', 'dt-num', fmt(r.proj, 'eur0')), el('td', 'dt-num', r.obj ? fmt(r.obj, 'eur0') : 'non saisi'));
            var tdJauge = el('td');
            if (r.pctProj !== null && r.pctProj !== undefined) {
                var jauge = el('div', 'tb-jauge');
                var rempli = el('div', 'tb-jauge-barre' + (r.pctProj >= 1 ? ' tb-jauge-ok' : ''));
                rempli.style.width = Math.min(100, r.pctProj * 100) + '%';
                jauge.append(rempli, el('span', null, fmt(r.pctProj, 'pct')));
                tdJauge.append(jauge);
            }
            tr.append(tdJauge, el('td', 'dt-num', fmt(r.pctReal, 'pct')));
            tbody.append(tr);
        });
        table.append(thead, tbody);
        zone.append(table);
        carte.append(zone, el('p', 'dt-muet tb-pied', b.el + ' jours ouvrés écoulés, ' + b.rest + ' restants. ' + (b.note || '')));
        return carte;
    }

    function dessinerBlocs(conteneur, blocs, ctx) {
        conteneur.textContent = '';
        (blocs || []).forEach(function (b) {
            if (b.xlOnly) return;
            var noeud = null;
            if (b.type === 'note') noeud = blocNote(b);
            else if (b.type === 'kpis') noeud = blocKpis(b);
            else if (b.type === 'table') noeud = blocTable(b, ctx);
            else if (b.type === 'chart') noeud = blocGraphique(b);
            else if (b.type === 'hbar') noeud = blocBarres(b);
            else if (b.type === 'heat') noeud = blocChaleur(b);
            else if (b.type === 'objectifs') noeud = blocObjectifs(b, ctx);
            if (noeud) conteneur.append(noeud);
        });
    }

    /* =====================================================================
       APPLICATION : filtres, requêtes, onglets, fiches, export
       Paramétrée par CONFIG (en tête du fichier)
       ===================================================================== */

    var etat = {
        admin: false,
        today: '',
        dMin: '',
        periode: 'mois',
        filtres: { du: '', au: '', pfs: [], canal: 'all', segs: [], clients: [], articles: [] },
        pfs: [],
        segs: [],
        donnees: null,
        onglet: CONFIG.onglets[0].k,
        fiches: [],
        numero: 0,
        minuterie: null,
        suivi: null,
        ctx: { tris: {}, pages: {}, tailles: {}, cleBase: '', admin: false, ouvrirFiche: null, saisirObjectifs: null }
    };
    var selecteurs = {};

    // ---------- Messages ----------
    function message(texte, type) {
        var zone = $('message');
        if (!texte) { afficher(zone, false); return; }
        dtMessage(zone, texte, type || 'info');
        afficher(zone, true);
    }

    // ---------- Périodes ----------
    function bornesPeriode(k) {
        var t = etat.today;
        var m0 = t.slice(0, 8) + '01';
        var du = t, au = t;
        if (k === 'hier') { du = ajouterJours(t, -1); au = du; }
        else if (k === '7j') du = ajouterJours(t, -6);
        else if (k === 'mois') du = m0;
        else if (k === 'moisp') { au = ajouterJours(m0, -1); du = au.slice(0, 8) + '01'; }
        else if (k === 'trim') { var q = Math.floor((+t.slice(5, 7) - 1) / 3) * 3 + 1; du = t.slice(0, 5) + String(q).padStart(2, '0') + '-01'; }
        else if (k === 'annee') du = t.slice(0, 4) + '-01-01';
        else if (k === '12m') du = (+t.slice(0, 4) - 1) + t.slice(4, 8) + '01';
        else if (k === 'tout') du = etat.dMin || '2025-01-01';
        else if (k === 'juin') du = '2026-06-01';
        return [du, au];
    }

    function choisirPeriode(k) {
        var b = bornesPeriode(k);
        etat.periode = k;
        $('du').value = b[0];
        $('au').value = b[1];
        dessinerPeriodes();
        programmer();
    }

    function dessinerPeriodes() {
        var zone = $('periodes');
        zone.textContent = '';
        CONFIG.periodes.forEach(function (p) {
            var actif = etat.periode === p.k;
            var b = el('button', 'tb-chip' + (actif ? ' tb-chip-actif' : ''), p.l);
            b.type = 'button';
            b.setAttribute('aria-pressed', actif ? 'true' : 'false');
            b.addEventListener('click', function () { choisirPeriode(p.k); });
            zone.append(b);
        });
    }

    // ---------- Puces plateformes / segments et canal ----------
    function dessinerPuces(zoneId, liste, choisis, libelle, basculer, toutVider) {
        var zone = $(zoneId);
        if (!zone) return;
        zone.textContent = '';
        var tout = el('button', 'tb-chip' + (choisis.length === 0 ? ' tb-chip-actif' : ''), 'Tout');
        tout.type = 'button';
        tout.addEventListener('click', toutVider);
        zone.append(tout);
        liste.forEach(function (valeur, i) {
            var cle = basculer.parIndex ? i : valeur;
            var actif = choisis.indexOf(cle) !== -1;
            var b = el('button', 'tb-chip' + (actif ? ' tb-chip-actif' : ''), libelle(valeur));
            b.type = 'button';
            b.setAttribute('aria-pressed', actif ? 'true' : 'false');
            b.addEventListener('click', function () { basculer(cle); });
            zone.append(b);
        });
    }

    function basculerDans(liste, cle, total) {
        var s = liste.slice();
        var j = s.indexOf(cle);
        if (j >= 0) s.splice(j, 1); else s.push(cle);
        if (s.length === total) return [];
        return s.sort();
    }

    function dessinerPlateformes() {
        var basculer = function (i) {
            etat.filtres.pfs = basculerDans(etat.filtres.pfs, i, etat.pfs.length);
            dessinerPlateformes();
            programmer();
        };
        basculer.parIndex = true;
        dessinerPuces('plateformes', etat.pfs, etat.filtres.pfs, function (p) { return p; }, basculer, function () {
            etat.filtres.pfs = [];
            dessinerPlateformes();
            programmer();
        });
    }

    function dessinerSegments() {
        var basculer = function (s) {
            etat.filtres.segs = basculerDans(etat.filtres.segs, s, etat.segs.length);
            dessinerSegments();
            programmer();
        };
        dessinerPuces('segments', etat.segs, etat.filtres.segs, function (s) { return s.charAt(0) + s.slice(1).toLowerCase(); }, basculer, function () {
            etat.filtres.segs = [];
            dessinerSegments();
            programmer();
        });
    }

    function dessinerCanal() {
        var zone = $('canal');
        if (!zone) return;
        zone.textContent = '';
        [['all', 'Tout'], ['pro', 'Pro'], ['part', 'Particulier']].forEach(function (c) {
            var actif = etat.filtres.canal === c[0];
            var b = el('button', 'tb-segment-choix' + (actif ? ' tb-segment-actif' : ''), c[1]);
            b.type = 'button';
            b.setAttribute('aria-pressed', actif ? 'true' : 'false');
            b.addEventListener('click', function () {
                etat.filtres.canal = c[0];
                dessinerCanal();
                programmer();
            });
            zone.append(b);
        });
    }

    function filtresCourants() {
        var f = {
            du: $('du').value,
            au: $('au').value,
            clients: selecteurs.clients.valeurs(),
            articles: selecteurs.articles.valeurs()
        };
        if (CONFIG.plateformes) { f.pfs = etat.filtres.pfs.slice(); f.canal = etat.filtres.canal; }
        if (CONFIG.segments) f.segs = etat.filtres.segs.slice();
        return f;
    }

    function programmer() {
        clearTimeout(etat.minuterie);
        etat.minuterie = setTimeout(lancerRequete, 300);
    }

    // ---------- État du chargement BC ----------
    function majEtat(e) {
        e = e || {};
        if (e.today) etat.today = e.today;
        if (e.dMin) etat.dMin = e.dMin;
        $('etat').textContent = e.loadedAt ? 'Données BC du ' + dateHeure(e.loadedAt) : 'Aucune donnée chargée';

        var bandeau = $('bandeau');
        if (e.erreurs && e.erreurs.length) {
            bandeau.textContent = e.erreurs.join(' · ');
            afficher(bandeau, true);
        } else {
            afficher(bandeau, false);
        }

        var zone = $('progression');
        if (e.run && e.run.on) {
            $('progression-barre').style.width = (e.run.pct || 0) + '%';
            $('progression-texte').textContent = e.run.text || 'Collecte BC en cours';
            afficher(zone, true);
            $('btn-recharger').disabled = true;
            if (!etat.suivi) etat.suivi = setInterval(suivreChargement, 5000);
        } else {
            afficher(zone, false);
            $('btn-recharger').disabled = false;
        }
    }

    async function suivreChargement() {
        try {
            var s = await dtFetch('statut');
            var avant = $('etat').textContent;
            majEtat(s.etat);
            if (!s.etat || !s.etat.run) {
                clearInterval(etat.suivi);
                etat.suivi = null;
                if ($('etat').textContent !== avant) message('Données Business Central rechargées.', 'info');
                lancerRequete();
            }
        } catch (e) {
            clearInterval(etat.suivi);
            etat.suivi = null;
            message(e.message, 'erreur');
        }
    }

    async function recharger() {
        $('btn-recharger').disabled = true;
        try {
            var r = await dtFetch('recharger', { body: {} });
            message(r.message || 'Chargement lancé', 'info');
            clearInterval(etat.suivi);
            etat.suivi = setInterval(suivreChargement, 5000);
        } catch (e) {
            $('btn-recharger').disabled = false;
            message(e.message, e.statut === 409 ? 'alerte' : 'erreur');
        }
    }

    // ---------- Requête principale ----------
    async function lancerRequete() {
        var numero = ++etat.numero;
        document.body.classList.add('tb-attente');
        try {
            var r = await dtFetch('requete', { body: { filtres: filtresCourants() }, timeout: 120000 });
            if (numero !== etat.numero) return;
            if (r.nodata) {
                etat.donnees = null;
                majEtat(r.etat);
                afficher($('vide'), true);
                afficher($('onglets'), false);
                $('blocs').textContent = '';
                return;
            }
            afficher($('vide'), false);
            afficher($('onglets'), true);
            etat.donnees = r;
            majEtat(r.etat);
            $('calcul').textContent = r.meta && r.meta.ms !== undefined ? 'calcul ' + r.meta.ms + ' ms' : '';
            selecteurs.clients.definirOptions((r.options.clients || []).map(function (o) { return { code: o.v, titre: o.t }; }));
            selecteurs.articles.definirOptions((r.options.articles || []).map(function (o) { return { code: o.v, titre: o.t }; }));
            if (etat.fiches.length) await rafraichirFiches();
            dessiner();
        } catch (e) {
            if (numero !== etat.numero) return;
            message(e.message, 'erreur');
        } finally {
            // le voile d'attente est toujours retiré, succès ou erreur
            if (numero === etat.numero) document.body.classList.remove('tb-attente');
        }
    }

    // ---------- Fiches article et client ----------
    async function chargerFiche(kind, code) {
        return dtFetch('fiche-' + kind, { body: { code: code, filtres: filtresCourants() }, timeout: 120000 });
    }

    async function ouvrirFiche(kind, code) {
        document.body.classList.add('tb-attente');
        try {
            var f = await chargerFiche(kind, code);
            etat.fiches.push({ kind: kind, code: code, titre: f.title, blocs: f.blocks });
            dessiner();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        } catch (e) {
            message(e.message, 'erreur');
        } finally {
            document.body.classList.remove('tb-attente');
        }
    }

    async function rafraichirFiches() {
        var nouvelles = [];
        for (var i = 0; i < etat.fiches.length; i++) {
            var x = etat.fiches[i];
            var f = await chargerFiche(x.kind, x.code);
            nouvelles.push({ kind: x.kind, code: x.code, titre: f.title, blocs: f.blocks });
        }
        etat.fiches = nouvelles;
    }

    // ---------- Rendu ----------
    function dessinerOnglets() {
        var zone = $('onglets');
        zone.textContent = '';
        CONFIG.onglets.forEach(function (o) {
            var actif = etat.onglet === o.k && etat.fiches.length === 0;
            var b = el('button', 'dt-onglet', o.l);
            b.type = 'button';
            b.setAttribute('role', 'tab');
            b.setAttribute('aria-selected', actif ? 'true' : 'false');
            b.addEventListener('click', function () {
                etat.onglet = o.k;
                etat.fiches = [];
                dessiner();
            });
            zone.append(b);
        });
    }

    function dessiner() {
        dessinerOnglets();
        var entete = $('fiche-entete');
        var blocs;
        if (etat.fiches.length) {
            var f = etat.fiches[etat.fiches.length - 1];
            $('fiche-type').textContent = f.kind === 'article' ? 'Fiche article' : 'Fiche client';
            $('fiche-titre').textContent = f.titre;
            afficher(entete, true);
            etat.ctx.cleBase = 'fiche' + etat.fiches.length + ':';
            blocs = f.blocs;
        } else {
            afficher(entete, false);
            etat.ctx.cleBase = etat.onglet + ':';
            blocs = etat.donnees && etat.donnees.tabs ? etat.donnees.tabs[etat.onglet] : [];
        }
        dessinerBlocs($('blocs'), blocs, etat.ctx);
    }

    // ---------- Objectifs (T&J) ----------
    function ouvrirObjectifs(bloc) {
        var dlg = $('dialogue-objectifs');
        if (!dlg) return;
        $('obj-mois').value = bloc.mois || etat.today.slice(0, 7);
        var zone = $('obj-champs');
        zone.textContent = '';
        var valeurs = {};
        bloc.rows.forEach(function (r) { if (!r.total && r.obj) valeurs[r.pf] = r.obj; });
        etat.pfs.forEach(function (pf) {
            var ligne = el('label', 'tb-obj-ligne');
            var input = el('input', 'dt-input tb-obj-valeur');
            input.type = 'number';
            input.min = '0';
            input.step = '100';
            input.dataset.pf = pf;
            input.value = valeurs[pf] !== undefined ? String(valeurs[pf]) : '';
            ligne.append(el('span', 'tb-obj-pf', pf), input, el('span', 'dt-muet', '€ HT'));
            zone.append(ligne);
        });
        afficher(dlg, true);
    }

    async function enregistrerObjectifs() {
        var vals = {};
        Array.prototype.forEach.call(document.querySelectorAll('.tb-obj-valeur'), function (input) {
            vals[input.dataset.pf] = input.value;
        });
        var bouton = $('obj-enregistrer');
        bouton.disabled = true;
        try {
            await dtFetch('objectifs', { body: { ym: $('obj-mois').value, vals: vals } });
            afficher($('dialogue-objectifs'), false);
            message('Objectifs enregistrés.', 'info');
            lancerRequete();
        } catch (e) {
            message(e.message, 'erreur');
        } finally {
            bouton.disabled = false;
        }
    }

    // ---------- Export ----------
    async function exporter() {
        var bouton = $('btn-export');
        var libelle = bouton.textContent;
        bouton.disabled = true;
        bouton.textContent = 'Export en cours…';
        try {
            var nom = await dtDownload('export', { filtres: filtresCourants() }, CONFIG.fichier);
            message('Fichier téléchargé : ' + nom, 'info');
        } catch (e) {
            message(e.message, 'erreur');
        } finally {
            bouton.disabled = false;
            bouton.textContent = libelle;
        }
    }

    // ---------- Démarrage ----------
    async function demarrer() {
        var utilisateur = await dtInit({ outil: CONFIG.outil, titre: CONFIG.titre });
        etat.admin = !!(utilisateur && utilisateur.admin);
        etat.ctx.admin = etat.admin;
        etat.ctx.ouvrirFiche = ouvrirFiche;
        etat.ctx.saisirObjectifs = CONFIG.objectifs ? ouvrirObjectifs : null;
        afficher($('btn-recharger'), etat.admin);

        var d = new Date();
        etat.today = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');

        selecteurs.clients = creerSelecteur($('sel-clients'), { libelle: 'Clients', onChange: programmer });
        selecteurs.articles = creerSelecteur($('sel-articles'), { libelle: 'Articles', onChange: programmer });

        $('du').addEventListener('change', function () { etat.periode = ''; dessinerPeriodes(); programmer(); });
        $('au').addEventListener('change', function () { etat.periode = ''; dessinerPeriodes(); programmer(); });
        $('btn-recharger').addEventListener('click', recharger);
        $('btn-export').addEventListener('click', exporter);
        $('btn-reinit').addEventListener('click', function () {
            etat.filtres.pfs = [];
            etat.filtres.segs = [];
            etat.filtres.canal = 'all';
            selecteurs.clients.vider();
            selecteurs.articles.vider();
            if (CONFIG.plateformes) { dessinerPlateformes(); dessinerCanal(); }
            if (CONFIG.segments) dessinerSegments();
            choisirPeriode('mois');
        });
        $('fiche-retour').addEventListener('click', function () { etat.fiches.pop(); dessiner(); });
        $('fiche-fermer').addEventListener('click', function () { etat.fiches = []; dessiner(); });
        if (CONFIG.objectifs) {
            $('obj-annuler').addEventListener('click', function () { afficher($('dialogue-objectifs'), false); });
            $('obj-enregistrer').addEventListener('click', enregistrerObjectifs);
        }

        try {
            var meta = await dtFetch('meta');
            majEtat(meta.etat);
            etat.pfs = meta.pfs || [];
            etat.segs = meta.segs || [];
            if (CONFIG.plateformes) { dessinerPlateformes(); dessinerCanal(); }
            if (CONFIG.segments) dessinerSegments();
        } catch (e) {
            message(e.message, 'erreur');
        }
        choisirPeriode('mois');
    }

    demarrer();

})();
