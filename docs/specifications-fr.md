# 1MANUEL D'UTILISATION, SPÉCIFICATION TECHNIQUE & CODE SOURCE DE RÉFÉRENCE

- **Version courante** : `1.1.0`
- **Date de mise à jour** : `2026-11-24`
- **Noyau d'inférence actif** : `whisperx`
- **Formats supportés** : `.mp4, .mkv, .mov, .avi, .mp3, .wav, .m4a, .flac, .ogg`
- **Catalogue de modèles** : `tiny, tiny.en, small, small.en, medium, medium.en, large-v3, large.en, pyannote, wav2vec2-fr`

SlideScript (v1.0.0) — Transcription, Diarisation & Extraction de Slides
Application Windows 100 % autonome sans installation : transcription vocale haute précision, identification automatique des intervenants, extraction sans perte (PNG) des diapositives, glisser-déposer direct dans la file d'attente, ajout dynamique en cours de lot, réordonnancement par glisser-déposer et export multi-format.

## 1. Objectif de l'Application & Guide d'Utilisation Rapide

SlideScript est une application Windows autonome conçue pour transformer automatiquement des enregistrements audio et vidéo (réunions, conférences, webinaires, formations, entretiens) en comptes rendus structurés, horodatés et attribués par intervenant, sans nécessiter aucune installation préalable ni connexion Internet lors des traitements.

### 1.1 À quoi sert l'application ? (Cas d'usage principaux)

Transcription fidèle & rapide : Convertit la parole en texte grâce aux modèles Whisper (CTranslate2), avec accélération automatique sur carte graphique NVIDIA CUDA ou adaptation dynamique aux cœurs CPU disponibles.
Identification des locuteurs (Diarisation) : Distingue et numérote automatiquement chaque participant (Speaker 1, Speaker 2, ...) et regroupe ses phrases consécutives avec leurs horodatages précis.
Capture haute définition et synchronisation des diapositives : Sur les vidéos de présentation (.mp4, .mkv, .mov, .avi), extrait les diapositives en PNG sans perte après stabilisation des animations et associe le discours de chaque intervenant à la slide correspondante.
Livrables prêts à diffuser : Génère simultanément, dans le même dossier que les fichiers sources, des fichiers Texte (.txt), Word (.docx), Word illustré avec captures d'écran (.docx) et PowerPoint (.pptx) contenant les notes du présentateur.

### 1.2 Comment utiliser l'application pas à pas ?

Étape 1 — Lancer l'exécutable portable : Double-cliquez sur .\SlideScript.exe depuis le dossier portable (disque local, clé USB ou lecteur réseau). L'application détecte automatiquement le matériel (GPU NVIDIA CUDA ou CPU) dans la barre supérieure sombre.
Étape 2 — Configuration initiale unique (si nécessaire) : Cliquez en haut à droite sur '⚙ Modèles installés & Mises à jour' pour installer, mettre à jour ou désinstaller en parallèle les modèles Whisper (tiny, small, medium, large-v3), le modèle d'identification des locuteurs (pyannote/speaker-diarization-community-1) et l'alignement français (wav2vec2-large-xlsr-53-french). Pour enregistrer votre jeton Hugging Face (hf_...), cliquez directement sur l'icône Hugging Face à droite de la ligne pyannote après avoir accepté les conditions des 2 dépôts requis.
Étape 3 — Sélectionner ou glisser-déposer les médias (Section 1 & Section 4) : Cliquez sur 'Choisir un fichier', 'Choisir plusieurs fichiers' ou 'Choisir un dossier' (cochez 'Inclure les sous-dossiers' pour traiter toute une arborescence), ou glissez-déposez directement vos fichiers/dossiers depuis l'Explorateur Windows dans la zone de la Section 4 — y compris pendant qu'un traitement est déjà en cours.
Étape 4 — Choisir les formats et paramètres (Section 2) : Cochez un ou plusieurs formats de sortie (.txt, .docx, .docx avec slides, .pptx avec transcripts dans les notes), sélectionnez la langue de l'audio (en, fr, auto, etc.) et le modèle Whisper souhaité. Toutes les transcriptions s'exécutent 100 % en local depuis .\hf_cache\.
Étape 5 — Démarrer et piloter le traitement en direct : Cliquez sur 'Démarrer la transcription'. Pendant l'exécution, les boutons 'Démarrer la transcription' et '⚙ Modèles installés & Mises à jour' sont désactivés pour éviter toute modification concurrente du cache local, tandis que vous pouvez suspendre le calcul ('Pause'), arrêter proprement le lot ('Arrêter'), ajouter de nouveaux fichiers par glisser-déposer dans la Section 4, réordonner les fichiers en attente via la poignée '≡', annuler/réactiver des fichiers (clic droit ou Suppr), modifier à chaud les cœurs CPU ou le modèle Whisper, et copier les lignes via Ctrl+C.

| Paramètre système | Spécification de référence |
| --- | --- |
| Exécutable & Version | .\SlideScript.exe — Version 1.0 bêta (Architecture v1.0.0 autoportante) |
| Environnement cible | Windows 10 / Windows 11 (x64) — Portable sans droits administrateur |
| Formats d'entrée | Vidéo : .mp4, .mkv, .mov, .avi  |  Audio : .mp3, .wav, .m4a, .flac, .ogg |
| Formats de sortie | TXT, DOCX standard, DOCX avec captures PNG haute netteté, PPTX avec notes du présentateur |
| Empreinte disque typique | ~4,2 Go (configuration anglaise allégée) à ~7,7 Go (avec alignement français Wav2Vec2 3,52 Go et modèles Whisper) |


## 2. Architecture Technique & Arborescence Relative

Le déploiement repose sur une arborescence relative 100 % autonome regroupant l'exécutable, les runtimes Python/PyInstaller, le cache de modèles et le dossier source (.\source\) à l'intérieur du même répertoire SlideScript :

| Chemin relatif (depuis l'exécutable) | Rôle et responsabilités techniques |
| --- | --- |
| .\SlideScript.exe | Exécutable principal compilé via PyInstaller (--onefile --noconsole) embarquant l'interface Tkinter/Canvas, le glisser-déposer Win32, l'icône Win32/Tk et les générateurs DOCX/PPTX. |
| .\runtime\ | Distribution Python 3.12.10 autonome complète contenant pip, PyInstaller, PyTorch, CTranslate2, WhisperX, pyannote.audio, huggingface_hub, Pillow, python-docx et python-pptx. |
| .\ffmpeg\ | Binaires autonomes ffmpeg.exe et ffprobe.exe injectés dynamiquement en tête du PATH système au lancement de l'application. |
| .\hf_cache\ | Cache local Hugging Face (HF_HUB_CACHE=.\hf_cache et TORCH_HOME=.\hf_cache\torch) hébergeant les modèles Faster-Whisper, Pyannote et Wav2Vec2. |
| .\run_engine_fast.py | Lanceur Python optimisé exécutant le pipeline WhisperX en sous-processus isolé avec sortie non bufferisée (PYTHONUNBUFFERED=1). |
| .\slidescript_config.json | Fichier JSON de persistance des préférences utilisateur créé automatiquement à côté de l'exécutable s'il est absent (langue UI, formats, langue audio, modèle, cœurs CPU, dernier dossier). |
| Registre Windows (HKCU) | Stockage sécurisé de la clé HF_TOKEN sous HKEY_CURRENT_USER\Software\SlideScript (DPAPI) (jamais écrite en clair dans .\slidescript_config.json). |
| .\source\SlideScript.py | Code source Python maître de l'interface graphique (v1.0.0). |
| .\source\SlideScript.ico | Icône source multi-résolution allégée (~8,15 Ko : BMP 16×16, 32×32 + PNG 48×48) embarquée dans l'exécutable lors de la compilation. |
| .\source\build.ps1 | Script PowerShell 100 % relatif invoquant .\runtime\python.exe -m PyInstaller pour recompiler .\SlideScript.exe sans dépendance externe. |


### 2.1 Gestion des icônes, certificats SSL d'entreprise et isolation PyInstaller

Double application d'icône Windows : L'icône SlideScript.ico (encodée en multi-résolution 16×16, 32×32 BMP et 48×48 PNG pour un poids optimal de 8,15 Ko) est embarquée via --icon et --add-data, puis appliquée de façon différée (self.after(150, ...)) via win.iconbitmap(), win.iconphoto() (PIL ImageTk) et les messages Win32 WM_SETICON (0x0080).
Redémarrage propre sans conflit _MEIPASS : Lors du redémarrage automatique après installation ou suppression d'un modèle, l'application purge les variables internes PyInstaller (_PYI_*, _MEIPASS2, TCL_LIBRARY, TK_LIBRARY) et déclenche la nouvelle instance via un processus détaché différé.

## 3. Spécifications de l'Interface Graphique (IHM) & Charte Visuelle

L'interface principale est structurée en une barre d'état matérielle sombre pleine largeur suivie de 5 sections fonctionnelles empilées verticalement, entièrement bilingues (Français / Anglais) avec bascule instantanée en temps réel (y compris sur l'étape en cours de traitement). Au démarrage, la zone Journal d'avancement adopte par défaut la même hauteur (6 lignes) que la file de la Section 4.

| Élément UI | Code couleur hexadécimal | Spécification visuelle et comportement |
| --- | --- | --- |
| Bandeau supérieur (top_bar) | #1E2229 (Ardoise très sombre) | Pleine largeur en haut de fenêtre : texte matériel blanc épuré sans emoji à gauche ; à droite, boutons 'Configurer token Hugging Face (OK)' (avec micro-logo Hugging Face N&B de 825 octets) et '⚙ Modèles installés & Mises à jour' harmonisés en gris clair (#E5E7EB, texte #111827 — passe en rouge #B91C1C / blanc si token manquant/invalide) et sélecteur FR/EN. |
| Barres de progression (PillProgressBar) | #0078B8 sur fond #DCE1E7 | Composant Canvas personnalisé à extrémités arrondies (pilule) utilisé dans la Section 3 (Avancement total et Fichier en cours avec pourcentage centré) ainsi que dans la fenêtre Modèles installés & Mises à jour. |
| Bouton principal (Démarrer) | #525864 (Gris foncé) + Icône #FFFFFF | Bouton plat avec icône triangle vectorielle blanche anti-aliasée (create_button_icons) assortie au texte blanc gras. |
| 4 Boutons secondaires | #D8DCE2 (Gris clair uniforme) + Icônes #111827 | Les 4 boutons à droite de Démarrer (Pause, Arrêter, Diagnostic système, Effacer le journal) partagent un fond gris clair uniforme (#D8DCE2) et des icônes vectorielles noires (#111827). |
| Zone Journal (txt_log) | #1E1E1E sur #E6E6E6 (height=6) | Hauteur initiale alignée sur la Section 4 (6 lignes visibles au démarrage) pour un équilibre visuel optimal entre la file d'attente et les logs. |


## 4. Gestion des Ressources Matérielles, Pilotage en Direct & File d'Attente

Détection matérielle GPU CUDA vs CPU : Au démarrage, detect_nvidia_cuda() vérifie nvidia-smi et torch.cuda.is_available(). En mode CUDA, --device cuda --compute_type float16 --batch_size 16 est utilisé et le sélecteur CPU est masqué. En mode CPU, --device cpu --compute_type int8 --batch_size 16 s'applique avec un sélecteur de cœurs (50 % à 100 % des cœurs logiques).
Glisser-déposer natif Windows dans la Section 4 : La fenêtre de la Section 4 (frame_queue et tree_queue) intercepte nativement les messages Windows WM_DROPFILES via ctypes/shell32.dll (enable_win32_drag_and_drop) : l'utilisateur peut glisser-déposer des fichiers ou des dossiers directement dans le tableau, avant ou pendant un traitement en cours.
Ajout dynamique de fichiers en cours de traitement : Les boutons de la Section 1 et le glisser-déposer de la Section 4 restent actifs pendant l'exécution d'un lot : les nouveaux médias sont ajoutés dynamiquement à la suite de la pile, sondés via ffprobe en arrière-plan, et pris en compte immédiatement dans la barre de progression totale et l'heure de fin estimée.
Réordonnancement dynamique (≡) & sélection haut de pile : La colonne de droite affiche la poignée '≡' permettant de glisser-déposer n'importe quelle ligne pour réorganiser l'ordre de passage. À chaque fin de fichier, le moteur sélectionne automatiquement le tout premier fichier non traité en partant du haut de la pile.
Annulation et réactivation de fichiers : Un clic droit (ou touche Suppr / Del) permet de basculer un fichier en statut 'Annulé' (ignoré lors de l'exécution et exclu du calcul de durée restante) ou de le réactiver à tout moment.
Ajustement dynamique des cœurs CPU et du modèle Whisper en cours de lot : Modifier le nombre de cœurs CPU applique immédiatement SetProcessAffinityMask sur le processus actif et met à jour les estimations restantes. Modifier le modèle Whisper en cours de traitement met à jour les colonnes Modèle et Tâches des fichiers en attente et s'applique automatiquement dès le fichier suivant.
Suivi temporel complet & heure de fin estimée : L'en-tête de la Section 4 affiche en continu le temps écoulé, le temps restant estimé et l'heure de fin prévisionnelle du lot (Fin estimée : JJ/MM HH:MM). La colonne 'Fin' affiche 'est.: AAAA-MM-JJ HH:MM:SS' pour le fichier en cours avant d'inscrire son horodatage réel une fois terminé.

## 5. Pipeline de Traitement, Diarisation & Extraction PNG Haute Définition


## 1. VAD & Transcription Faster-Whisper : Sélection automatique de la variante <modèle>.en lorsque la langue est 'en' et que le modèle spécialisé est présent dans .\hf_cache\.


## 2. Alignement & Diarisation Pyannote : Alignement phonétique au niveau des mots (wav2vec2-large-xlsr-53-french pour le français) et identification des locuteurs via pyannote/speaker-diarization-community-1 et segmentation-3.0.


## 3. Extraction sans perte (PNG) des diapositives vidéo : FFmpeg extrait une image toutes les 2 secondes en PNG sans perte (-vf fps=0.5,scale=iw:ih:flags=lanczos). La détection compare la zone centrale utile (5 %–95 % largeur, 6 %–85 % hauteur) via Pillow et retient la frame confirmée après 3 images stables (6 s), garantissant une capture nette après la fin des animations PowerPoint, insérée à 6,3 pouces dans Word et en plein écran 16:9 dans PowerPoint.


| Format de sortie | Nom avec locuteurs cochés | Nom sans locuteurs | Structure et contenu du livrable |
| --- | --- | --- | --- |
| Texte brut (txt) | <nom>_<modèle>_speakers.txt | <nom>_<modèle>.txt | Transcription horodatée par bloc de locuteur avec sauts de ligne par phrase. |
| Word standard (docx) | <nom>_<modèle>_speakers.docx | <nom>_<modèle>.docx | Document mis en page (police Aptos), pastilles colorées distinctes par locuteur et horodatage. |
| Word avec slides (docx) | <nom>_<modèle>_speakers_slides.docx | <nom>_<modèle>_slides.docx | Capture PNG sans perte (largeur 6,3") de chaque diapositive suivie des interventions associées. |
| PowerPoint (pptx) | <nom>_<modèle>_speakers_slides.pptx | <nom>_<modèle>_slides.pptx | Présentation 16:9 avec chaque diapositive PNG en plein écran et la transcription en notes du présentateur. |


## 6. Guide de Modification du Code & Procédure de Recompilation Autoportante

L'application est conçue pour être maintenue et recompilée de manière 100 % autoportante sur n'importe quel poste Windows, sans aucune dépendance à un environnement Python externe ni à un dossier projet séparé : l'interpréteur console Python 3.12.10, pip et PyInstaller sont intégrés directement dans .\runtime\.

### 6.1 Comment modifier le code source (.\source\SlideScript.py)


## 1. Édition du code : Ouvrez .\source\SlideScript.py dans un éditeur UTF-8 (VS Code, Notepad++, Bloc-notes). Enregistrez toujours en UTF-8 sans BOM.


## 2. Repères dans le fichier : Les libellés bilingues sont centralisés dans le dictionnaire I18N ('fr' et 'en'). La logique de file d'attente, le glisser-déposer Win32 (enable_win32_drag_and_drop), l'ajout dynamique (append_files_to_queue) et l'interface se trouvent dans WhisperXApp.


## 3. Contrôle syntaxique préalable : Avant toute recompilation, vérifiez la validité syntaxique du fichier avec le Python embarqué : .\runtime\python.exe -m py_compile .\source\SlideScript.py


## 4. Ajout éventuel d'un module Python : Pour installer une bibliothèque Python supplémentaire dans le runtime portable : .\runtime\python.exe -m pip install --no-warn-script-location <nom_paquet>


### 6.2 Comment recompiler l'exécutable en 1 commande

Ouvrez un terminal PowerShell dans le dossier racine SlideScript (ou faites un clic droit sur .\source\build.ps1 → Exécuter avec PowerShell) et lancez :

| powershell -ExecutionPolicy Bypass -File .\source\build.ps1 |
| --- |


| Paramètre | Valeur / Résolution relative |
| --- | --- |
| $PortableRoot | Répertoire contenant .\SlideScript.exe (parent de .\source\). |
| $SourceDir | Répertoire des sources : Join-Path $PortableRoot 'source' (soit .\source\). |
| $RuntimePy | Interpréteur autonome embarqué : Join-Path $PortableRoot 'runtime\python.exe'. |
| $GuiFile | Code source Python maître : Join-Path $SourceDir 'SlideScript.py'. |
| $BuildIcon | Icône source multi-résolution (8,15 Ko) : Join-Path $SourceDir 'SlideScript.ico'. |
| Commande de build | powershell -ExecutionPolicy Bypass -File .\source\build.ps1 |


### 6.3 Contenu intégral de .\source\build.ps1 (recompilation 100 % autoportante)


| $SourceDir    = $PSScriptRoot $PortableRoot = Split-Path $SourceDir -Parent $RuntimePy    = Join-Path $PortableRoot 'runtime\python.exe' $GuiFile      = Join-Path $SourceDir 'SlideScript.py' $BuildIcon    = Join-Path $SourceDir 'SlideScript.ico' $BuildTemp    = Join-Path $PortableRoot 'build_temp'  Get-Process -Name 'WhisperX_Transcription' -ErrorAction SilentlyContinue | Stop-Process -Force Start-Sleep -Milliseconds 300 Remove-Item -Force (Join-Path $PortableRoot 'SlideScript.exe') -ErrorAction SilentlyContinue  if (Test-Path $BuildTemp) { Remove-Item -Recurse -Force -LiteralPath $BuildTemp } New-Item -ItemType Directory -Path $BuildTemp -Force | Out-Null  $ArgsList = @(     '-m', 'PyInstaller',     '--clean', '--noconsole', '--onefile',     '--collect-all', 'pptx',     '--collect-all', 'lxml',     '--hidden-import=docx',     '--hidden-import=PIL',     '--hidden-import=PIL.Image',     '--hidden-import=PIL.ImageChops',     '--hidden-import=PIL.ImageDraw',     '--hidden-import=PIL.ImageStat',     '--hidden-import=PIL.ImageTk',     '--name', 'WhisperX_Transcription',     '--distpath', $PortableRoot,     '--workpath', (Join-Path $BuildTemp 'work'),     '--specpath', (Join-Path $BuildTemp 'spec') ) if (Test-Path $BuildIcon) { $ArgsList += @("--icon=$BuildIcon", "--add-data=$BuildIcon;.") } $ArgsList += $GuiFile  & $RuntimePy @ArgsList Remove-Item -Recurse -Force -LiteralPath $BuildTemp -ErrorAction SilentlyContinue |
| --- |


### 6.4 Gestionnaire de Modèles Parallèle, Certificats SSL d'Entreprise & Sauvegardes Automatiques (v1.0.0)

- Opérations parallèles & télémétrie en direct : La fenêtre '⚙ Modèles installés & Mises à jour' permet de lancer simultanément des installations, mises à jour et désinstallations (Whisper, Pyannote et Wav2Vec2 français) via des sous-processus non bloquants supervisés par la boucle Tkinter, avec barre de progression, débit en Mo/s, temps restant estimé et bouton 'Arrêter' individuel.
- Détection automatique des mises à jour distantes : Dès l'ouverture de la fenêtre, un contrôle asynchrone léger via HfApi compare l'empreinte locale (refs/main) avec Hugging Face ; si une révision plus récente existe, le bouton 'Mettre à jour' passe automatiquement en bleu ('Mise à jour dispo').
- Validation SSL stricte en environnement d'entreprise (win_ca_bundle.pem) : Les certificats publics des autorités racines approuvées par Windows (Cert:\LocalMachine\Root et Cert:\CurrentUser\Root) sont exportés dans .\source\win_ca_bundle.pem et déclarés via REQUESTS_CA_BUNDLE et SSL_CERT_FILE, assurant le fonctionnement derrière un proxy d'inspection TLS sans jamais désactiver la vérification SSL.
- Verrouillage pendant la transcription active : Dès le clic sur 'Démarrer la transcription', les boutons 'Démarrer la transcription' et '⚙ Modèles installés & Mises à jour' deviennent non cliquables (state=DISABLED) jusqu'à la fin ou l'arrêt du lot, garantissant l'intégrité du cache local pendant l'exécution de WhisperX.
- Sauvegardes horodatées & restauration automatique : Avant chaque recompilation de .\SlideScript.exe, le script .\source\build.ps1 archive l'exécutable précédent dans .\backups\WhisperX_Transcription_AAAAMMJJ_HHMMSS.exe (conservation des 5 dernières versions) et le restaure automatiquement en cas d'échec PyInstaller.