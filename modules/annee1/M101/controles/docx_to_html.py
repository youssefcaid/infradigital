#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_to_html.py
================
Convertit un fichier Word (.docx) en une page HTML qui respecte exactement
le même gabarit que le reste du site "Infrastructure Digitale" (navbar,
mode sombre, breadcrumb, style des cours/exercices/contrôles/EFM), PUIS
ajoute automatiquement une carte vers cette page dans le index.html de la
sous-section (cours / exercices / controles / efm).

INSTALLATION (une seule fois) :
    pip install python-docx

UTILISATION :
    python3 docx_to_html.py "modules/annee1/M101/controles/controle1.docx"

Le script devine automatiquement :
    - l'année      (annee1 / annee2)   -> depuis le chemin du dossier
    - le module    (M101, M203, ...)   -> depuis le chemin du dossier
    - la sous-section (cours/exercices/controles/efm) -> depuis le nom du dossier parent
    - le numéro    (1, 2, 3...)        -> depuis les chiffres dans le nom du fichier

Le fichier .docx DOIT donc se trouver dans le bon dossier, par exemple :
    modules/annee1/M101/controles/controle1.docx
    modules/annee2/M205/efm/efm4.docx
    modules/annee1/M107/cours/cours2.docx

Si le nom ou l'emplacement ne permettent pas de deviner un élément,
utilisez les options --annee, --module, --section, --num, --titre.

Exemples :
    python3 docx_to_html.py modules/annee1/M101/controles/controle1.docx
    python3 docx_to_html.py mon_document.docx --annee annee1 --module M101 --section controles --num 1 --titre "Contrôle sur les réseaux"
"""

import argparse
import html
import os
import re
import shutil
import sys

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("ERREUR : le module 'python-docx' n'est pas installé.")
    print("Installez-le avec :  pip install python-docx")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Données du site (mêmes informations que le générateur du site)
# ---------------------------------------------------------------------------
ANNEE_LABELS = {"annee1": "1ère année", "annee2": "2ème année"}

MODULES = {
    "EGTS106": "Culture et techniques avancées du numérique",
    "M101": "Métier et formation",
    "M102": "Les enjeux d'un système d'information",
    "M103": "Conception d'un réseau informatique",
    "M104": "Fonctionnement d'un système d'exploitation",
    "M105": "Gestion de l'infrastructure virtualisée",
    "M106": "Automatisation des tâches d'administration",
    "M107": "Sécurité des systèmes d'information",
    "M108": "Processus et outils de veille technologique",
    "M201": "Mise en place d'une infrastructure réseaux",
    "M202": "Administration d'un environnement Windows",
    "M203": "Administration d'un environnement Linux",
    "M204": "Enjeux de la technologie SDN",
    "M205": "Administration d'un environnement Cloud",
    "M206": "Sécurité d'une infrastructure digitale",
}

# slug de dossier -> (label de la section, label au singulier pour une page)
SECTIONS = {
    "cours":     ("Cours",      "Cours"),
    "exercices": ("Exercices",  "Exercice"),
    "controles": ("Contrôles",  "Contrôle"),
    "efm":       ("EFM",        "EFM"),
}

# ---------------------------------------------------------------------------
# CSS partagé (identique à celui utilisé pour générer tout le site)
# ---------------------------------------------------------------------------
SHARED_CSS = """
:root {
    --primary: #3b82f6; --primary-dark: #1d4ed8; --secondary: #10b981;
    --accent: #f59e0b; --dark: #1f2937; --light: #f9fafb; --indigo: #4f46e5; --nav-h: 68px;
}
html { background: #f9fafb; }
.dark { color-scheme: dark; }
html.dark { background: #050a18; }
body { font-family: 'Inter', sans-serif; transition: background-color .3s, color .3s; }
.card-hover { transition: transform .3s ease, box-shadow .3s ease; }
.card-hover:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,.1); }
.dark .card-hover:hover { box-shadow: 0 5px 15px rgba(0,0,0,.3); }
.animate-fadeIn { animation: fadeIn .6s ease-out; }
@keyframes fadeIn { from { opacity:0; transform:translateY(10px);} to {opacity:1; transform:translateY(0);} }

.navbar { position: sticky; top: 0; z-index: 100; height: var(--nav-h); display:flex; align-items:center;
    background: rgba(255,255,255,0.78); backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
    border-bottom: 1px solid rgba(15,23,42,0.08); transition: background-color .3s, border-color .3s, box-shadow .3s; }
.dark .navbar { background: rgba(5,10,24,0.55); border-bottom-color: rgba(0,150,255,0.12); }
.navbar.scrolled { box-shadow: 0 10px 26px -16px rgba(15,23,42,0.35); }
.dark .navbar.scrolled { background: rgba(5,10,24,0.85); border-bottom-color: rgba(0,180,255,0.22); box-shadow: 0 8px 30px rgba(0,10,30,0.5); }
.nav-inner { width:100%; max-width:1180px; margin:0 auto; padding:0 20px; display:flex; align-items:center; justify-content:space-between; gap:16px; }
.brand { display:flex; align-items:center; gap:9px; font-weight:700; text-decoration:none; white-space:nowrap; }
.brand-mark { width:34px; height:34px; flex-shrink:0; display:flex; align-items:center; justify-content:center; }
.brand-mark img { width:100%; height:100%; object-fit:contain; display:block; }
.brand-name { font-size:1rem; letter-spacing:-.01em; color:#0f172a; }
.dark .brand-name { color:#e8ecff; }
.brand-name strong { color: var(--primary); }
.dark .brand-name strong { background: linear-gradient(135deg,#ffe066,#ffc800 60%,#ff9500); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; color:transparent; }
.nav-links { display:flex; align-items:center; gap:30px; margin-left:auto; margin-right:22px; }
.nav-links a { font-weight:500; font-size:.9rem; color:#475569; text-decoration:none; position:relative; padding:6px 0; transition:color .2s ease; }
.nav-links a::after { content:''; position:absolute; left:0; bottom:0; width:0; height:2px; border-radius:2px; background:linear-gradient(90deg,var(--primary),var(--indigo)); transition:width .25s ease, background .3s ease; }
.nav-links a:hover { color:#0f172a; }
.nav-links a:hover::after, .nav-links a.active::after { width:100%; }
.nav-links a.active { color:var(--primary-dark); font-weight:600; }
.dark .nav-links a { color:#8899cc; }
.dark .nav-links a:hover { color:#fff; }
.dark .nav-links a::after { background:linear-gradient(90deg,#00c8ff,#ffc800); }
.dark .nav-links a.active { color:#00c8ff; }
.nav-actions { display:flex; align-items:center; gap:10px; flex-shrink:0; }
.theme-toggle { width:38px; height:38px; border-radius:10px; border:1.5px solid rgba(59,130,246,0.25); background:rgba(59,130,246,0.06); color:var(--primary-dark);
    display:flex; align-items:center; justify-content:center; cursor:pointer; position:relative; transition:background-color .2s, border-color .2s, color .3s; }
.theme-toggle:hover { background:rgba(59,130,246,0.14); }
.dark .theme-toggle { border-color:rgba(0,180,255,0.3); background:rgba(0,144,255,0.08); color:#00c8ff; }
.dark .theme-toggle:hover { background:rgba(0,170,255,0.18); }
.theme-toggle svg { width:17px; height:17px; position:absolute; transition:opacity .25s, transform .3s; }
.theme-toggle .icon-sun { opacity:0; transform:rotate(-40deg) scale(.6); }
.theme-toggle .icon-moon { opacity:1; transform:rotate(0) scale(1); }
.dark .theme-toggle .icon-sun { opacity:1; transform:rotate(0) scale(1); }
.dark .theme-toggle .icon-moon { opacity:0; transform:rotate(40deg) scale(.6); }
.burger { display:none; width:38px; height:38px; border-radius:10px; border:1.5px solid rgba(59,130,246,0.25); background:rgba(59,130,246,0.06);
    align-items:center; justify-content:center; flex-direction:column; gap:5px; cursor:pointer; flex-shrink:0; }
.dark .burger { border-color:rgba(0,180,255,0.28); background:rgba(0,144,255,0.08); }
.burger span { width:18px; height:2px; border-radius:2px; background:var(--primary-dark); transition:all .25s ease, background .3s ease; }
.dark .burger span { background:#00c8ff; }
.burger[aria-expanded="true"] span:nth-child(1) { transform:translateY(7px) rotate(45deg); }
.burger[aria-expanded="true"] span:nth-child(2) { opacity:0; }
.burger[aria-expanded="true"] span:nth-child(3) { transform:translateY(-7px) rotate(-45deg); }
.mobile-menu { position:absolute; top:100%; left:0; right:0; z-index:99; background:rgba(255,255,255,0.97); backdrop-filter:blur(16px); -webkit-backdrop-filter:blur(16px);
    border-bottom:1px solid rgba(15,23,42,0.08); display:flex; flex-direction:column; padding:8px 20px 20px;
    transform:translateY(-8px); opacity:0; pointer-events:none; transition:transform .25s ease, opacity .25s ease; box-shadow:0 16px 30px -18px rgba(15,23,42,0.25); }
.mobile-menu.is-open { transform:translateY(0); opacity:1; pointer-events:auto; }
.dark .mobile-menu { background:rgba(6,12,28,0.97); border-bottom-color:rgba(0,180,255,0.2); box-shadow:0 16px 30px -18px rgba(0,0,0,0.6); }
.mobile-menu a { font-weight:500; font-size:.95rem; color:#334155; text-decoration:none; padding:14px 4px; border-bottom:1px solid rgba(15,23,42,0.06); }
.mobile-menu a:last-child { border-bottom:none; }
.mobile-menu a.active { color:var(--primary-dark); font-weight:600; }
.dark .mobile-menu a { color:#e8ecff; border-bottom-color:rgba(0,150,255,0.1); }
.dark .mobile-menu a.active { color:#00c8ff; }
.mobile-menu-label { font-size:.72rem; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:#94a3b8; padding:14px 4px 4px; }
.dark .mobile-menu-label { color:#5b6b96; }
@media (min-width:769px) { .mobile-menu { display:none !important; } }
@media (max-width:768px) { .nav-links { display:none; } .burger { display:flex; } }

.nav-drop { position:relative; }
.nav-drop > button { font-weight:500; font-size:.9rem; color:#475569; background:none; border:none; cursor:pointer;
    display:flex; align-items:center; gap:5px; padding:6px 0; font-family:inherit; line-height:1; }
.nav-drop > button:hover { color:#0f172a; }
.dark .nav-drop > button { color:#8899cc; }
.dark .nav-drop > button:hover { color:#fff; }
.nav-drop svg { width:12px; height:12px; transition:transform .2s ease; }
.nav-drop.is-open > button svg { transform:rotate(180deg); }
.nav-drop-menu { position:absolute; top:calc(100% + 14px); left:50%; transform:translateX(-50%) translateY(-6px);
    min-width:190px; background:rgba(255,255,255,0.98); backdrop-filter:blur(14px); -webkit-backdrop-filter:blur(14px);
    border:1px solid rgba(15,23,42,0.08); border-radius:12px; padding:8px; box-shadow:0 16px 30px -14px rgba(15,23,42,0.25);
    opacity:0; pointer-events:none; transition:opacity .2s ease, transform .2s ease; }
.nav-drop.is-open .nav-drop-menu { opacity:1; pointer-events:auto; transform:translateX(-50%) translateY(0); }
.nav-drop:hover .nav-drop-menu, .nav-drop:focus-within .nav-drop-menu { opacity:1; pointer-events:auto; transform:translateX(-50%) translateY(0); }
.nav-drop:hover > button svg, .nav-drop:focus-within > button svg { transform:rotate(180deg); }
.dark .nav-drop-menu { background:rgba(6,12,28,0.98); border-color:rgba(0,180,255,0.2); box-shadow:0 16px 30px -14px rgba(0,0,0,0.55); }
.nav-drop-menu a { display:flex; align-items:center; gap:8px; padding:9px 12px; border-radius:8px; font-size:.88rem; font-weight:500; color:#334155; text-decoration:none; }
.nav-drop-menu a:hover { background:rgba(59,130,246,0.08); color:#1d4ed8; }
.dark .nav-drop-menu a { color:#c9d3f0; }
.dark .nav-drop-menu a:hover { background:rgba(0,144,255,0.12); color:#00c8ff; }

.contenu-redige { line-height:1.85; }
.contenu-redige h2 { font-size:1.35rem; font-weight:800; margin-top:2rem; margin-bottom:.75rem; color:#1d4ed8; }
.dark .contenu-redige h2 { color:#5aa9ff; }
.contenu-redige h3 { font-size:1.1rem; font-weight:700; margin-top:1.5rem; margin-bottom:.5rem; color:#1f2937; }
.dark .contenu-redige h3 { color:#e8ecff; }
.contenu-redige p { margin-bottom:1rem; color:#374151; }
.dark .contenu-redige p { color:#c3cbe6; }
.contenu-redige ul, .contenu-redige ol { margin:0 0 1rem 1.4rem; color:#374151; }
.dark .contenu-redige ul, .dark .contenu-redige ol { color:#c3cbe6; }
.contenu-redige li { margin-bottom:.4rem; }
.contenu-redige code { background:rgba(59,130,246,0.1); color:#1d4ed8; padding:2px 6px; border-radius:4px; font-size:.9em; }
.dark .contenu-redige code { background:rgba(0,144,255,0.15); color:#5aa9ff; }
.contenu-redige pre { background:#0f172a; color:#e2e8f0; padding:1rem; border-radius:10px; overflow-x:auto; margin-bottom:1rem; }
.contenu-redige pre code { background:none; color:inherit; padding:0; }
.contenu-redige blockquote { border-inline-start:4px solid var(--primary); padding:.5rem 1rem; background:rgba(59,130,246,0.06); border-radius:0 8px 8px 0; margin-bottom:1rem; color:#374151; }
.dark .contenu-redige blockquote { color:#c3cbe6; background:rgba(0,144,255,0.08); }
.contenu-redige table { width:100%; border-collapse:collapse; margin-bottom:1.25rem; font-size:.92rem; }
.contenu-redige th, .contenu-redige td { border:1px solid #d1d5db; padding:.55rem .75rem; text-align:start; color:#374151; }
.dark .contenu-redige th, .dark .contenu-redige td { border-color:#374151; color:#c3cbe6; }
.contenu-redige th { background:rgba(59,130,246,0.08); font-weight:700; color:#1d4ed8; }
.dark .contenu-redige th { background:rgba(0,144,255,0.12); color:#5aa9ff; }
"""

NAV_SCRIPT = """
(function () {
    var navbar = document.getElementById('navbar');
    var burgerBtn = document.getElementById('burgerBtn');
    var mobileMenu = document.getElementById('mobileMenu');
    var themeToggle = document.getElementById('themeToggle');
    var drop = document.getElementById('modulesDrop');

    window.addEventListener('scroll', function () {
        navbar.classList.toggle('scrolled', window.scrollY > 8);
    }, { passive: true });

    function applyTheme(isDark) {
        document.documentElement.classList.toggle('dark', isDark);
        localStorage.setItem('theme', isDark ? 'dark' : 'light');
        themeToggle.setAttribute('aria-label', isDark ? 'Activer le thème clair' : 'Activer le thème sombre');
    }
    var saved = localStorage.getItem('theme');
    applyTheme(saved ? saved === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches);
    themeToggle.addEventListener('click', function () { applyTheme(!document.documentElement.classList.contains('dark')); });

    function openMenu() {
        burgerBtn.setAttribute('aria-expanded', 'true');
        mobileMenu.classList.add('is-open');
        document.addEventListener('keydown', onKeydown);
        document.addEventListener('click', onOutsideClick, true);
    }
    function closeMenu() {
        burgerBtn.setAttribute('aria-expanded', 'false');
        mobileMenu.classList.remove('is-open');
        document.removeEventListener('keydown', onKeydown);
        document.removeEventListener('click', onOutsideClick, true);
    }
    function onKeydown(e) { if (e.key === 'Escape') closeMenu(); }
    function onOutsideClick(e) { if (!mobileMenu.contains(e.target) && !burgerBtn.contains(e.target)) closeMenu(); }
    burgerBtn.addEventListener('click', function () {
        burgerBtn.getAttribute('aria-expanded') === 'true' ? closeMenu() : openMenu();
    });

    if (drop) {
        var dbtn = drop.querySelector('button');
        function closeDrop() { drop.classList.remove('is-open'); dbtn.setAttribute('aria-expanded', 'false'); }
        function toggleDrop() {
            var open = drop.classList.toggle('is-open');
            dbtn.setAttribute('aria-expanded', open ? 'true' : 'false');
        }
        dbtn.addEventListener('click', function (e) { e.stopPropagation(); toggleDrop(); });
        document.addEventListener('click', function (e) { if (!drop.contains(e.target)) closeDrop(); });
        document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDrop(); });
    }
})();
"""


def navbar_html(root):
    return f'''    <header class="navbar" id="navbar">
        <div class="nav-inner">
            <a href="{root}index.html" class="brand">
                <span class="brand-mark"><img src="{root}logo.png" alt="Infrastructure Digitale"></span>
                <span class="brand-name">Infrastructure&nbsp;<strong>Digitale</strong></span>
            </a>
            <nav class="nav-links" id="navLinks">
                <div class="nav-drop" id="modulesDrop">
                    <button type="button" aria-haspopup="true" aria-expanded="false">
                        Modules
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
                    </button>
                    <div class="nav-drop-menu">
                        <a href="{root}modules/annee1/index.html"><i class="fas fa-book-reader"></i> 1ère année</a>
                        <a href="{root}modules/annee2/index.html"><i class="fas fa-laptop-code"></i> 2ème année</a>
                    </div>
                </div>
                <a href="{root}applications.html">Applications</a>
                <a href="{root}apropos.html">À propos</a>
                <a href="{root}confidentialite.html">Confidentialité</a>
                <a href="{root}contact.html">Contact</a>
            </nav>
            <div class="nav-actions">
                <button id="themeToggle" type="button" class="theme-toggle" aria-label="Activer le thème sombre">
                    <svg class="icon-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8.5 8.5 0 1 1 9.5 4a6.6 6.6 0 0 0 10.5 10.5Z"/></svg>
                    <svg class="icon-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2.5M12 19.5V22M4.93 4.93l1.77 1.77M17.3 17.3l1.77 1.77M2 12h2.5M19.5 12H22M4.93 19.07l1.77-1.77M17.3 6.7l1.77-1.77"/></svg>
                </button>
                <button class="burger" id="burgerBtn" type="button" aria-label="Ouvrir le menu" aria-expanded="false" aria-controls="mobileMenu">
                    <span></span><span></span><span></span>
                </button>
            </div>
        </div>
        <div class="mobile-menu" id="mobileMenu">
            <span class="mobile-menu-label">Modules</span>
            <a href="{root}modules/annee1/index.html"><i class="fas fa-book-reader mr-2"></i>1ère année</a>
            <a href="{root}modules/annee2/index.html"><i class="fas fa-laptop-code mr-2"></i>2ème année</a>
            <a href="{root}applications.html">Applications</a>
            <a href="{root}apropos.html">À propos</a>
            <a href="{root}confidentialite.html">Confidentialité</a>
            <a href="{root}contact.html">Contact</a>
        </div>
    </header>
'''


def footer_html(root):
    return '''    <footer class="mt-8 border-t border-gray-200 dark:border-gray-700 py-6">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex flex-col md:flex-row justify-between items-center">
                <div class="mb-4 md:mb-0 text-center md:text-left">
                    <h3 class="text-lg font-bold text-blue-600 dark:text-blue-400">Infrastructure Digitale</h3>
                    <p class="text-sm text-gray-500 dark:text-gray-400">Ressources pédagogiques pour étudiants OFPPT</p>
                </div>
                <div class="flex space-x-4 mb-4 md:mb-0">
                    <a href="https://facebook.com/youssefelcaid1" target="_blank" class="text-gray-500 hover:text-blue-600 dark:hover:text-blue-400" aria-label="Facebook"><i class="fab fa-facebook text-xl"></i></a>
                    <a href="https://instagram.com/youssefelcaid1" target="_blank" class="text-gray-500 hover:text-pink-600 dark:hover:text-pink-400" aria-label="Instagram"><i class="fab fa-instagram text-xl"></i></a>
                    <a href="https://youtube.com/professeuryoussef1" target="_blank" class="text-gray-500 hover:text-red-600 dark:hover:text-red-400" aria-label="YouTube"><i class="fab fa-youtube text-xl"></i></a>
                    <a href="https://wa.me/212708482810" target="_blank" class="text-gray-500 hover:text-green-600 dark:hover:text-green-400" aria-label="WhatsApp"><i class="fab fa-whatsapp text-xl"></i></a>
                </div>
            </div>
            <div class="text-center mt-4 pt-4 border-t border-gray-200 dark:border-gray-700">
                <p class="text-xs text-gray-500 dark:text-gray-400">&copy; <span id="currentYear"></span> Infrastructure Digitale – Professeur Youssef. Tous droits réservés.</p>
            </div>
        </div>
    </footer>
'''


def crumb(items):
    parts = []
    for i, (label, href) in enumerate(items):
        if i > 0:
            parts.append('            <i class="fas fa-chevron-right text-[10px] mx-1 opacity-50"></i>')
        if href:
            parts.append(f'            <a href="{href}" class="hover:text-blue-600 dark:hover:text-blue-400 transition">{label}</a>')
        else:
            parts.append(f'            <span class="text-gray-700 dark:text-gray-300 font-medium">{label}</span>')
    return "\n".join(parts)


def page(root, title, description, breadcrumb, body):
    return f'''<!DOCTYPE html>
<html lang="fr" dir="ltr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <meta name="description" content="{description}">
    <script async src="https://www.googletagmanager.com/gtag/js?id=G-1TMFJG7WBV"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){{dataLayer.push(arguments);}}
      gtag('js', new Date());
      gtag('config', 'G-1TMFJG7WBV');
    </script>
    <link rel="icon" type="image/png" href="{root}logo.png">
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {{
            darkMode: 'class',
            theme: {{ extend: {{ colors: {{ primary: '#3b82f6', 'primary-dark': '#1d4ed8', secondary: '#10b981', accent: '#f59e0b' }} }} }}
        }}
    </script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@100..900&display=swap" rel="stylesheet">
    <style>{SHARED_CSS}</style>
</head>
<body class="text-gray-900 dark:text-gray-100 min-h-screen flex flex-col">
{navbar_html(root)}
    <main class="flex-grow max-w-5xl mx-auto px-3 sm:px-6 lg:px-8 py-6 sm:py-10 w-full">
        <nav class="text-sm text-gray-500 dark:text-gray-400 mb-6 flex flex-wrap items-center gap-1" aria-label="Fil d'Ariane">
{breadcrumb}
        </nav>
{body}
    </main>
{footer_html(root)}
    <script>
        document.getElementById('currentYear').textContent = new Date().getFullYear();
        {NAV_SCRIPT}
    </script>
</body>
</html>
'''


def content_page(annee_key, annee_label, code, nom, slug, item_label, num, titre, root, article_html):
    body = f'''
        <div class="animate-fadeIn">
            <div class="mb-6">
                <span class="inline-block px-3 py-1 rounded-full bg-blue-100 dark:bg-blue-800 text-blue-600 dark:text-blue-300 text-xs font-bold mb-3">{item_label} {num}</span>
                <h1 class="text-3xl font-extrabold mb-2">{titre}</h1>
                <p class="text-gray-500 dark:text-gray-400">Module {code} — {nom}</p>
            </div>
            <article class="contenu-redige bg-white dark:bg-gray-800 rounded-2xl p-6 sm:p-8 shadow-lg border border-gray-200 dark:border-gray-700">
{article_html}
            </article>
            <div class="mt-8 flex justify-between items-center text-sm">
                <a href="index.html" class="text-blue-600 dark:text-blue-400 hover:underline font-medium">
                    <i class="fas fa-arrow-left mr-1"></i> Retour à la liste « {SECTIONS_SLUG_TO_SECTION_LABEL(slug)} »
                </a>
            </div>
        </div>
'''
    section_label = SECTIONS[slug][0]
    return page(
        root, f"{titre} – {item_label} {num} – {code} – Infrastructure Digitale",
        f"{titre} — {item_label} {num} du module {code} ({nom}).",
        crumb([("Accueil", root + "index.html"), ("Modules", root + "modules/index.html"),
               (annee_label, root + f"modules/{annee_key}/index.html"),
               (code, root + f"modules/{annee_key}/{code}/index.html"),
               (section_label, root + f"modules/{annee_key}/{code}/{slug}/index.html"),
               (f"{item_label} {num}", None)]),
        body
    )


def SECTIONS_SLUG_TO_SECTION_LABEL(slug):
    return SECTIONS[slug][0]


def new_card_html(slug, item_label, num, titre):
    return (f'                <a href="{slug}{num}.html" class="card-hover bg-white dark:bg-gray-800 rounded-xl p-5 '
            f'shadow-md border border-gray-200 dark:border-gray-700 flex items-center justify-between">\n'
            f'                    <div class="flex items-center gap-3">\n'
            f'                        <span class="w-9 h-9 flex items-center justify-center rounded-lg bg-blue-100 '
            f'dark:bg-blue-800 text-blue-600 dark:text-blue-300 font-bold text-sm">{num}</span>\n'
            f'                        <span class="font-semibold text-gray-800 dark:text-gray-100">{item_label} {num} — {html.escape(titre)}</span>\n'
            f'                    </div>\n'
            f'                    <i class="fas fa-chevron-right text-gray-400 dark:text-gray-500"></i>\n'
            f'                </a>')


# ---------------------------------------------------------------------------
# Lecture et conversion du fichier Word
# ---------------------------------------------------------------------------
def iter_block_items(document):
    """Parcourt le document dans l'ordre réel : paragraphes ET tableaux mélangés."""
    parent_elm = document.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, document)
        elif child.tag == qn('w:tbl'):
            yield Table(child, document)


def run_to_html(run):
    text = html.escape(run.text or "")
    if not text:
        return ""
    if run.bold:
        text = f"<strong>{text}</strong>"
    if run.italic:
        text = f"<em>{text}</em>"
    if run.underline:
        text = f"<u>{text}</u>"
    return text


def paragraph_to_inline_html(paragraph):
    return "".join(run_to_html(r) for r in paragraph.runs) or html.escape(paragraph.text or "")


def table_to_html(table):
    rows_html = []
    for i, row in enumerate(table.rows):
        cells_html = []
        tag = "th" if i == 0 else "td"
        for cell in row.cells:
            text = html.escape(cell.text.strip())
            cells_html.append(f"<{tag}>{text}</{tag}>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
    return "                <table>\n                    " + "\n                    ".join(rows_html) + "\n                </table>"


def docx_to_article_html(path):
    document = docx.Document(path)
    html_parts = []
    list_buffer = []
    list_tag = None
    title_candidate = None

    def flush_list():
        nonlocal list_buffer, list_tag
        if list_buffer:
            items = "\n".join(f"                    <li>{t}</li>" for t in list_buffer)
            html_parts.append(f"                <{list_tag}>\n{items}\n                </{list_tag}>")
            list_buffer = []
            list_tag = None

    for block in iter_block_items(document):
        if isinstance(block, Table):
            flush_list()
            html_parts.append(table_to_html(block))
            continue

        para = block
        text = para.text.strip()
        style = (para.style.name or "").lower()

        if not text:
            flush_list()
            continue

        inline = paragraph_to_inline_html(para)

        if "heading 1" in style or "titre 1" in style or style == "title":
            flush_list()
            if title_candidate is None:
                # Le tout premier titre "Heading 1" sert de titre de page (h1) :
                # on ne le réécrit pas une deuxième fois dans le corps de l'article.
                title_candidate = text
                continue
            html_parts.append(f"                <h2>{inline}</h2>")
        elif "heading 2" in style or "titre 2" in style:
            flush_list()
            html_parts.append(f"                <h2>{inline}</h2>")
        elif "heading" in style or "titre" in style:
            flush_list()
            html_parts.append(f"                <h3>{inline}</h3>")
        elif "list bullet" in style or "liste à puces" in style or text.startswith(("- ", "• ", "● ")):
            if list_tag != "ul":
                flush_list()
                list_tag = "ul"
            cleaned = re.sub(r"^[-•●]\s*", "", inline)
            list_buffer.append(cleaned)
        elif "list number" in style or "liste numéros" in style or re.match(r"^\d+[.)]\s", text):
            if list_tag != "ol":
                flush_list()
                list_tag = "ol"
            cleaned = re.sub(r"^\d+[.)]\s*", "", inline)
            list_buffer.append(cleaned)
        else:
            flush_list()
            html_parts.append(f"                <p>{inline}</p>")

    flush_list()
    return "\n".join(html_parts), title_candidate


# ---------------------------------------------------------------------------
# Mise à jour du index.html de la sous-section
# ---------------------------------------------------------------------------
def update_section_index(index_path, slug, item_label, num, titre):
    with open(index_path, "r", encoding="utf-8") as f:
        content = f.read()

    backup_path = index_path + ".bak"
    if not os.path.exists(backup_path):
        shutil.copy2(index_path, backup_path)

    # Numéro déjà présent ?
    existing_nums = re.findall(rf'href="{re.escape(slug)}(\d+)\.html"', content)
    if str(num) in existing_nums:
        print(f"  -> ATTENTION : une carte pour {slug}{num}.html existe déjà dans {index_path} ; "
              f"aucune carte ajoutée (fichier html quand même régénéré).")
        return

    # Supprimer le bloc "aucun contenu publié" s'il existe
    empty_pattern = re.compile(
        r'\s*<div class="text-center py-14[^>]*>.*?</div>\s*\n',
        re.DOTALL
    )
    content = empty_pattern.sub("\n", content, count=1)

    # Insérer la nouvelle carte juste avant la fermeture du conteneur "space-y-3"
    marker = '<div class="space-y-3">'
    pos_open = content.find(marker)
    if pos_open == -1:
        print(f"  -> ATTENTION : impossible de trouver le conteneur des cartes dans {index_path}. "
              f"Ajoutez la carte manuellement.")
        return
    pos_close = content.find("</div>", pos_open)
    if pos_close == -1:
        print(f"  -> ATTENTION : structure inattendue dans {index_path}. Ajoutez la carte manuellement.")
        return

    card = new_card_html(slug, item_label, num, titre)
    content = content[:pos_close] + card + "\n            " + content[pos_close:]

    with open(index_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  -> Carte ajoutée dans {index_path}")


# ---------------------------------------------------------------------------
# Détection année / module / section / numéro depuis le chemin
# ---------------------------------------------------------------------------
def guess_from_path(docx_path):
    parts = [p for p in os.path.normpath(docx_path).split(os.sep) if p]
    annee_key = module_code = slug = None
    for p in parts:
        if p in ("annee1", "annee2"):
            annee_key = p
        elif p in MODULES:
            module_code = p
        elif p in SECTIONS:
            slug = p
    filename = os.path.splitext(os.path.basename(docx_path))[0]
    m = re.search(r"(\d+)\s*$", filename)
    num = m.group(1) if m else None
    return annee_key, module_code, slug, num


def compute_root(docx_path):
    """Calcule le préfixe relatif vers la racine du site (ex: '../../../../')."""
    abs_path = os.path.normpath(docx_path)
    parts = abs_path.split(os.sep)
    if "modules" in parts:
        idx = parts.index("modules")
        # nombre de dossiers entre la racine du site et le dossier du fichier
        # (on exclut le nom du fichier lui-même, dernier élément de "parts")
        depth = len(parts) - idx - 1  # 'modules' + annee + code + slug = 4 en général
        depth = max(depth, 1)
        return "../" * depth
    # repli : suppose la structure standard modules/anneeX/CODE/slug/fichier.docx
    return "../../../../"


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Convertit un .docx en page HTML du site et met à jour le index.html correspondant.")
    parser.add_argument("docx_path", help="Chemin vers le fichier .docx à convertir")
    parser.add_argument("--annee", choices=["annee1", "annee2"], help="Forcer l'année (sinon devinée depuis le chemin)")
    parser.add_argument("--module", help="Forcer le code du module, ex: M101 (sinon deviné depuis le chemin)")
    parser.add_argument("--section", choices=list(SECTIONS.keys()), help="Forcer la sous-section (sinon devinée depuis le chemin)")
    parser.add_argument("--num", help="Forcer le numéro de l'élément, ex: 2 (sinon deviné depuis le nom du fichier)")
    parser.add_argument("--titre", help="Forcer le titre affiché (sinon le premier titre du document Word est utilisé)")
    args = parser.parse_args()

    docx_path = args.docx_path
    if not os.path.isfile(docx_path):
        print(f"ERREUR : fichier introuvable : {docx_path}")
        sys.exit(1)

    guessed_annee, guessed_module, guessed_slug, guessed_num = guess_from_path(docx_path)
    annee_key = args.annee or guessed_annee
    module_code = args.module or guessed_module
    slug = args.section or guessed_slug
    num = args.num or guessed_num

    missing = []
    if not annee_key:
        missing.append("--annee (annee1 ou annee2)")
    if not module_code:
        missing.append("--module (ex: M101)")
    if not slug:
        missing.append("--section (cours, exercices, controles ou efm)")
    if not num:
        missing.append("--num (ex: 1)")
    if missing:
        print("ERREUR : impossible de deviner automatiquement les informations suivantes :")
        for m in missing:
            print("   ", m)
        print("\nPlacez le fichier dans le bon dossier (ex: modules/annee1/M101/controles/controle1.docx)")
        print("ou précisez les options manquantes en ligne de commande.")
        sys.exit(1)

    if module_code not in MODULES:
        print(f"ATTENTION : le module '{module_code}' n'est pas dans la liste connue ; "
              f"le nom du module ne sera pas affiché correctement. Vous pouvez éditer le script "
              f"pour l'ajouter dans le dictionnaire MODULES.")
        module_nom = ""
    else:
        module_nom = MODULES[module_code]

    annee_label = ANNEE_LABELS[annee_key]
    section_label, item_label = SECTIONS[slug]

    print(f"Conversion : {docx_path}")
    print(f"  Année        : {annee_label} ({annee_key})")
    print(f"  Module       : {module_code} — {module_nom}")
    print(f"  Section      : {section_label} ({slug})")
    print(f"  Numéro       : {num}")

    article_html, title_from_doc = docx_to_article_html(docx_path)
    titre = args.titre or title_from_doc or f"{item_label} {num}"
    print(f"  Titre        : {titre}")

    root = compute_root(docx_path)
    page_html = content_page(annee_key, annee_label, module_code, module_nom, slug, item_label, num, titre, root, article_html)

    out_dir = os.path.dirname(os.path.abspath(docx_path))
    out_path = os.path.join(out_dir, f"{slug}{num}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page_html)
    print(f"  -> Page créée : {out_path}")

    index_path = os.path.join(out_dir, "index.html")
    if os.path.isfile(index_path):
        update_section_index(index_path, slug, item_label, num, titre)
    else:
        print(f"  -> ATTENTION : {index_path} introuvable, la carte n'a pas pu être ajoutée automatiquement.")

    print("\nTerminé. Vous pouvez ouvrir le fichier .html dans un navigateur pour vérifier le résultat.")


if __name__ == "__main__":
    main()
