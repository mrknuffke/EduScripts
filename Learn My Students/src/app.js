/**
 * Learn My Students - App Logic
 */

(function () {
  'use strict';

  // --- STATE ---
  let loadedDecks = []; // Array of deck objects
  let activeSession = null; // Current active study session state
  let currentImportPending = null; // Deck waiting for import confirmation
  let hasUnsavedChanges = false;
  let activeGridDeck = null; // Current deck displayed in Class Grid view
  let gridNamesHidden = true; // Whether names are hidden in grid view
  let gridRevealedIds = new Set(); // Card IDs individually revealed in grid view
  let editingCardId = null; // Card ID currently in preferred name inline edit mode

  // --- BROWSER STORAGE PERSISTENCE (IndexedDB + localStorage) ---
  const DB_NAME = 'LearnMyStudentsDB';
  const STORE_NAME = 'decks_store';
  const DB_VERSION = 1;

  function openDB() {
    return new Promise((resolve) => {
      if (!window.indexedDB) {
        resolve(null);
        return;
      }
      try {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME);
          }
        };
        request.onsuccess = (e) => resolve(e.target.result);
        request.onerror = (e) => {
          console.warn('IndexedDB error:', e.target.error);
          resolve(null);
        };
      } catch (err) {
        console.warn('IndexedDB exception:', err);
        resolve(null);
      }
    });
  }

  async function persistDecksToStorage() {
    try {
      const db = await openDB();
      if (db) {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        store.put(loadedDecks, 'loaded_decks_data');
        await new Promise((res) => {
          tx.oncomplete = res;
          tx.onerror = res;
        });
      } else {
        localStorage.setItem('learn_my_students_decks', JSON.stringify(loadedDecks));
      }
    } catch (err) {
      console.warn('Failed to persist decks via IndexedDB, trying localStorage fallback:', err);
      try {
        localStorage.setItem('learn_my_students_decks', JSON.stringify(loadedDecks));
      } catch (e) {
        console.warn('LocalStorage fallback failed:', e);
      }
    }
  }

  async function loadDecksFromStorage() {
    try {
      const db = await openDB();
      if (db) {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const req = store.get('loaded_decks_data');
        const data = await new Promise((res) => {
          req.onsuccess = () => res(req.result);
          req.onerror = () => res(null);
        });
        if (data && Array.isArray(data) && data.length > 0) {
          return data;
        }
      }
      const localData = localStorage.getItem('learn_my_students_decks');
      if (localData) {
        const parsed = JSON.parse(localData);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed;
        }
      }
    } catch (err) {
      console.warn('Failed to load decks from storage:', err);
    }
    return null;
  }

  // Matrix multiplication helper: A * B
  function multiplyMatrix(m1, m2) {
    return [
      m1[0] * m2[0] + m1[2] * m2[1],
      m1[1] * m2[0] + m1[3] * m2[1],
      m1[0] * m2[2] + m1[2] * m2[3],
      m1[1] * m2[2] + m1[3] * m2[3],
      m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
      m1[1] * m2[4] + m1[3] * m2[5] + m1[5]
    ];
  }

  // Set up PDF.js worker
  function initPdfWorker() {
    if (typeof pdfjsLib === 'undefined') {
      console.error('PDF.js library not loaded.');
      return;
    }

    try {
      if (typeof PDF_WORKER_SRC !== 'undefined' && PDF_WORKER_SRC.length > 0) {
        const blob = new Blob([PDF_WORKER_SRC], { type: 'text/javascript' });
        pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
      } else {
        pdfjsLib.GlobalWorkerOptions.workerSrc = '';
      }
    } catch (err) {
      console.warn('Failed to initialize PDF worker Blob URL, falling back to main thread fake worker:', err);
      pdfjsLib.GlobalWorkerOptions.workerSrc = '';
    }
  }

  // --- PDF PARSER ---
  async function parsePDF(arrayBuffer, fileName) {
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdfDoc = await loadingTask.promise;
    const extractedCards = [];
    let parseWarnings = [];

    const defaultDeckName = fileName.replace(/\.pdf$/i, '').replace(/_/g, ' ');

    for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: 3.0 }); // 3x scale for crisp photo cropping

      // 1. Extract image rectangles using operator list
      const opList = await page.getOperatorList();
      const imageRects = [];
      let ctmStack = [[1, 0, 0, 1, 0, 0]];

      for (let i = 0; i < opList.fnArray.length; i++) {
        const fn = opList.fnArray[i];
        const args = opList.argsArray[i];

        if (fn === pdfjsLib.OPS.save) {
          ctmStack.push([...ctmStack[ctmStack.length - 1]]);
        } else if (fn === pdfjsLib.OPS.restore) {
          if (ctmStack.length > 1) ctmStack.pop();
        } else if (fn === pdfjsLib.OPS.transform) {
          const currentCTM = ctmStack[ctmStack.length - 1];
          ctmStack[ctmStack.length - 1] = multiplyMatrix(currentCTM, args);
        } else if (fn === pdfjsLib.OPS.paintImageXObject || fn === pdfjsLib.OPS.paintJpegXObject) {
          const currentCTM = ctmStack[ctmStack.length - 1];
          const [a, b, c, d, e, f] = currentCTM;
          const p1 = [e, f];
          const p2 = [a + e, b + f];
          const p3 = [c + e, d + f];
          const p4 = [a + c + e, b + d + f];

          const x0 = Math.min(p1[0], p2[0], p3[0], p4[0]);
          const x1 = Math.max(p1[0], p2[0], p3[0], p4[0]);
          const y0 = Math.min(p1[1], p2[1], p3[1], p4[1]);
          const y1 = Math.max(p1[1], p2[1], p3[1], p4[1]);

          // Filter out tiny images (like icons or decor)
          if ((x1 - x0) > 30 && (y1 - y0) > 30) {
            imageRects.push({ x0, y0, x1, y1, ctm: currentCTM });
          }
        }
      }

      // Sort imageRects in reading order (top-to-bottom, left-to-right)
      imageRects.sort((a, b) => {
        if (Math.abs(b.y0 - a.y0) > 20) return b.y0 - a.y0; // Higher y in PDF coords is earlier row
        return a.x0 - b.x0; // Left to right
      });

      // 2. Extract text items
      const textContent = await page.getTextContent();
      const rawItems = textContent.items
        .filter(it => it.str && it.str.trim())
        .map(it => ({
          text: it.str.trim(),
          x0: it.transform[4],
          x1: it.transform[4] + it.width,
          y: it.transform[5]
        }));

      // 3. Render page to canvas for cropping photos
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport }).promise;

      // 4. Pair images with text items (§4.1)
      for (let imgIdx = 0; imgIdx < imageRects.length; imgIdx++) {
        const rect = imageRects[imgIdx];

        // In PDF page coords (bottom-left origin), rect.y0 is bottom edge of photo
        const candidates = rawItems.filter(item =>
          item.y <= (rect.y0 + 2) &&
          item.y >= (rect.y0 - 45) &&
          item.x1 > (rect.x0 - 15) &&
          item.x0 < (rect.x1 + 15)
        );

        // Sort candidates top-to-bottom (higher y to lower y in PDF coords)
        candidates.sort((a, b) => b.y - a.y);

        let studentName = null;
        let warningState = null;

        if (candidates.length === 0) {
          warningState = 'No text lines found beneath photo';
        } else {
          // Top baseline item(s) form the student name
          const topY = candidates[0].y;
          const nameItems = candidates.filter(c => Math.abs(c.y - topY) < 2);
          nameItems.sort((a, b) => a.x0 - b.x0);
          studentName = nameItems.map(c => c.text).join(' ');
          // Note: any candidate item below topY is the student ID, which is DISCARDED & NEVER STORED.
        }

        // Crop photo from rendered canvas (§4.2)
        const pTL = viewport.convertToViewportPoint(rect.x0, rect.y1);
        const pBR = viewport.convertToViewportPoint(rect.x1, rect.y0);

        const cropX = Math.max(0, Math.min(pTL[0], pBR[0]));
        const cropY = Math.max(0, Math.min(pTL[1], pBR[1]));
        const cropW = Math.min(canvas.width - cropX, Math.abs(pBR[0] - pTL[0]));
        const cropH = Math.min(canvas.height - cropY, Math.abs(pBR[1] - pTL[1]));

        let photoDataUrl = '';
        if (cropW > 0 && cropH > 0) {
          const cropCanvas = document.createElement('canvas');
          // Cap long edge at 400px (§4.2)
          const scaleDown = Math.min(1.0, 400 / Math.max(cropW, cropH));
          cropCanvas.width = Math.round(cropW * scaleDown);
          cropCanvas.height = Math.round(cropH * scaleDown);
          const cropCtx = cropCanvas.getContext('2d');
          cropCtx.drawImage(canvas, cropX, cropY, cropW, cropH, 0, 0, cropCanvas.width, cropCanvas.height);
          photoDataUrl = cropCanvas.toDataURL('image/jpeg', 0.82);
        }

        extractedCards.push({
          id: 'c_' + Math.random().toString(36).substr(2, 9),
          name: studentName || 'Unknown Student',
          preferredName: null,
          photo: photoDataUrl,
          box: 1,
          dueSession: 1,
          seen: 0,
          correct: 0,
          missed: 0,
          warning: warningState
        });
      }
    }

    // Check for duplicate names (claimed by two photos)
    const nameCounts = new Map();
    for (const card of extractedCards) {
      if (card.name !== 'Unknown Student') {
        nameCounts.set(card.name, (nameCounts.get(card.name) || 0) + 1);
      }
    }
    for (const card of extractedCards) {
      if (nameCounts.get(card.name) > 1) {
        card.warning = 'Name claimed by multiple photos';
        parseWarnings.push(`Duplicate name found: "${card.name}"`);
      }
    }

    return {
      deckName: defaultDeckName,
      cards: extractedCards,
      warnings: parseWarnings
    };
  }

  // --- NORMALISATION & LEVENSHTEIN (§6) ---
  function normalizeString(str) {
    if (!str) return '';
    return str
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Strip combining marks
      .toLowerCase()
      .replace(/['’'‘\-‑‒–—]/g, ' ') // Replace hyphens and apostrophes with spaces
      .replace(/[^\w\s]/g, '') // Strip all remaining punctuation
      .replace(/\s+/g, ' ') // Collapse whitespace
      .trim();
  }

  // --- DATE & TIME FORMATTERS ---
  function formatRelativeTime(isoString) {
    if (!isoString) return 'Not yet run';
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return 'Not yet run';

    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    if (diffMs < 0) return 'Just now';

    const diffSecs = Math.floor(diffMs / 1000);
    const diffMins = Math.floor(diffSecs / 60);
    const diffHours = Math.floor(diffMins / 60);
    const diffDays = Math.floor(diffHours / 24);

    if (diffSecs < 45) return 'Just now';
    if (diffMins < 60) return `${diffMins}m ago`;
    if (diffHours < 24) return `${diffHours}h ago`;
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays}d ago`;

    return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function formatFullDateTime(isoString) {
    if (!isoString) return 'Never run';
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return 'Never run';
    return date.toLocaleString(undefined, {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: 'numeric',
      minute: '2-digit'
    });
  }

  function levenshtein(a, b) {
    const matrix = [];
    for (let i = 0; i <= b.length; i++) matrix[i] = [i];
    for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    return matrix[b.length][a.length];
  }

  function allowedEdits(targetLen) {
    if (targetLen <= 3) return 0;
    if (targetLen <= 7) return 1;
    return 2;
  }

  function gradeAnswer(inputStr, card) {
    const normInput = normalizeString(inputStr);
    if (!normInput) return { outcome: 'miss', targetSpelling: card.name };

    const targetFull = card.preferredName || card.name;
    const normTargetFull = normalizeString(targetFull);
    const normRosterFull = normalizeString(card.name);

    // 1. Full name match check (against preferred or roster name)
    const distFull = levenshtein(normInput, normTargetFull);
    const distRoster = levenshtein(normInput, normRosterFull);

    if (distFull <= allowedEdits(normTargetFull.length) || distRoster <= allowedEdits(normRosterFull.length)) {
      return { outcome: 'correct', targetSpelling: card.name };
    }

    // 2. Single token match check
    const tokens = normTargetFull.split(' ').concat(normRosterFull.split(' '));
    const uniqueTokens = Array.from(new Set(tokens.filter(t => t.length > 0)));

    for (const token of uniqueTokens) {
      const distToken = levenshtein(normInput, token);
      if (distToken <= allowedEdits(token.length)) {
        return { outcome: 'partial', targetSpelling: card.name };
      }
    }

    // Multi-token input comparison check
    const inputTokens = normInput.split(' ');
    if (inputTokens.length > 1) {
      // Check if all input tokens match some token in target
      let allTokensMatched = true;
      for (const inTok of inputTokens) {
        let matched = false;
        for (const tok of uniqueTokens) {
          if (levenshtein(inTok, tok) <= allowedEdits(tok.length)) {
            matched = true;
            break;
          }
        }
        if (!matched) {
          allTokensMatched = false;
          break;
        }
      }
      if (allTokensMatched) {
        return { outcome: 'partial', targetSpelling: card.name };
      }
    }

    return { outcome: 'miss', targetSpelling: card.name };
  }

  // --- SCHEDULER (LEITNER §7) ---
  const BOX_INTERVALS = [0, 1, 3, 6, 12, 25]; // 1-indexed by box number

  function startStudySession(decksToStudy, isDrillEverything = false) {
    // Increment sessionCount for each involved deck
    decksToStudy.forEach(deck => {
      deck.sessionCount = (deck.sessionCount || 0) + 1;
      deck.updatedAt = new Date().toISOString();
    });
    setUnsaved(true);

    let candidateCards = [];
    decksToStudy.forEach(deck => {
      deck.cards.forEach(card => {
        candidateCards.push({ card, deck });
      });
    });

    let sessionQueue = [];

    if (isDrillEverything) {
      // Review every card in random order
      sessionQueue = candidateCards.sort(() => Math.random() - 0.5);
    } else {
      // Due cards: dueSession <= deck.sessionCount
      const dueCandidates = candidateCards.filter(item => item.card.dueSession <= item.deck.sessionCount);

      // Group by box, lowest box first, shuffle within box
      const boxGroups = { 1: [], 2: [], 3: [], 4: [], 5: [] };
      dueCandidates.forEach(item => {
        const b = Math.min(Math.max(1, item.card.box || 1), 5);
        boxGroups[b].push(item);
      });

      for (let b = 1; b <= 5; b++) {
        boxGroups[b].sort(() => Math.random() - 0.5);
        sessionQueue = sessionQueue.concat(boxGroups[b]);
      }

      // Allow full class roster study sessions (up to 50 cards)
      sessionQueue = sessionQueue.slice(0, 50);
    }

    activeSession = {
      decks: decksToStudy,
      queue: sessionQueue,
      currentIndex: 0,
      isDrillEverything,
      stats: { correct: 0, partial: 0, missed: 0 },
      missedCards: [],
      deckStatsMap: new Map()
    };

    return activeSession;
  }

  function recordCardResult(outcome) {
    if (!activeSession || activeSession.currentIndex >= activeSession.queue.length) return;

    const currentItem = activeSession.queue[activeSession.currentIndex];
    const { card, deck } = currentItem;

    if (!activeSession.deckStatsMap) {
      activeSession.deckStatsMap = new Map();
    }
    let dStats = activeSession.deckStatsMap.get(deck);
    if (!dStats) {
      dStats = { correct: 0, partial: 0, missed: 0, total: 0 };
      activeSession.deckStatsMap.set(deck, dStats);
    }
    dStats.total++;

    card.seen = (card.seen || 0) + 1;

    if (outcome === 'correct') {
      card.correct = (card.correct || 0) + 1;
      card.box = Math.min((card.box || 1) + 1, 5);
      card.dueSession = deck.sessionCount + BOX_INTERVALS[card.box];
      activeSession.stats.correct++;
      dStats.correct++;
    } else if (outcome === 'partial') {
      // Box unchanged
      card.dueSession = deck.sessionCount + BOX_INTERVALS[card.box || 1];
      activeSession.stats.partial++;
      dStats.partial++;
    } else {
      // Miss
      card.missed = (card.missed || 0) + 1;
      card.box = 1;
      card.dueSession = deck.sessionCount + 1;
      activeSession.stats.missed++;
      dStats.missed++;
      if (!activeSession.missedCards.includes(currentItem)) {
        activeSession.missedCards.push(currentItem);
      }
    }

    deck.updatedAt = new Date().toISOString();
    setUnsaved(true);
  }

  // --- UI RENDERERS & EVENT HANDLERS ---

  function setUnsaved(val) {
    hasUnsavedChanges = val;
    const badge = document.getElementById('unsaved-badge');
    if (badge) {
      if (hasUnsavedChanges) badge.classList.remove('hidden');
      else badge.classList.add('hidden');
    }
  }

  function showScreen(screenId) {
    document.body.className = 'screen-' + screenId;
  }

  function renderHome() {
    const listEl = document.getElementById('deck-list');
    listEl.innerHTML = '';

    const exportCollectiveHomeBtn = document.getElementById('btn-export-collective-home');
    if (exportCollectiveHomeBtn) {
      exportCollectiveHomeBtn.disabled = loadedDecks.length === 0;
    }

    if (loadedDecks.length === 0) {
      listEl.innerHTML = '<div class="empty-state">No decks loaded yet. Drop a PDF or deck file above to get started!</div>';
      document.getElementById('btn-study-combined').disabled = true;
      return;
    }

    document.getElementById('btn-study-combined').disabled = loadedDecks.length < 2;

    loadedDecks.forEach((deck, idx) => {
      const totalCards = deck.cards ? deck.cards.length : 0;
      const dueCount = deck.cards ? deck.cards.filter(c => (c.dueSession || 1) <= ((deck.sessionCount || 0) + 1)).length : 0;
      const box45Count = deck.cards ? deck.cards.filter(c => (c.box || 1) >= 4).length : 0;
      const masteryPct = totalCards > 0 ? Math.round((box45Count / totalCards) * 100) : 0;

      const lastRunIso = deck.lastStudiedAt || (deck.sessionCount > 0 ? deck.updatedAt : null);
      const relativeRunTime = formatRelativeTime(lastRunIso);
      const fullRunTime = formatFullDateTime(lastRunIso);

      let scoreBadgeHtml = '';
      if (deck.lastScore !== undefined && deck.lastScore !== null) {
        const score = Math.round(deck.lastScore);
        let badgeClass = 'score-badge';
        let icon = '🌱';
        if (score >= 95) {
          badgeClass += ' distinction';
          icon = '👑';
        } else if (score >= 80) {
          badgeClass += ' meeting';
          icon = '⭐';
        } else if (score >= 55) {
          badgeClass += ' developing';
          icon = '🚀';
        } else {
          badgeClass += ' emerging';
          icon = '🌱';
        }
        let detailText = '';
        if (deck.lastRunStats && deck.lastRunStats.total > 0) {
          detailText = ` (${deck.lastRunStats.correct}/${deck.lastRunStats.total} correct)`;
        }
        scoreBadgeHtml = `<span class="${badgeClass}" title="Last Session Score: ${score}%${escapeHtml(detailText)}">${icon} Last Score: ${score}%</span>`;
      } else {
        scoreBadgeHtml = `<span class="score-badge empty" title="No completed study session score recorded yet">No score yet</span>`;
      }

      const itemEl = document.createElement('div');
      itemEl.className = 'deck-item';
      itemEl.innerHTML = `
        <div class="deck-info">
          <input type="checkbox" class="deck-checkbox" data-index="${idx}">
          <div class="deck-details-wrapper">
            <div class="deck-header-row">
              <span class="deck-title">${escapeHtml(deck.deckName)}</span>
              ${scoreBadgeHtml}
            </div>
            <div class="deck-indicators">
              <span class="indicator-pill" title="Last study run: ${escapeHtml(fullRunTime)}">
                <span>🕒</span> Last run: <strong>${escapeHtml(relativeRunTime)}</strong>
              </span>
              <span class="indicator-pill" title="${dueCount} of ${totalCards} cards currently due for review">
                <span>👥</span> <strong>${totalCards}</strong> students (<span class="due-text">${dueCount} due</span>)
              </span>
              <span class="indicator-pill" title="${box45Count} of ${totalCards} students in Box 4 or Box 5 (mastered)">
                <span>🏆</span> Mastery: <strong>${masteryPct}%</strong> (${box45Count}/${totalCards} Box 4+)
              </span>
            </div>
          </div>
        </div>
        <div class="deck-actions">
          <button class="btn primary btn-study-deck" data-index="${idx}">Study (${dueCount} due)</button>
          <button class="btn secondary btn-drill-deck" data-index="${idx}">Drill All</button>
          <button class="btn secondary btn-grid-deck" data-index="${idx}">Class Grid</button>
        </div>
      `;
      listEl.appendChild(itemEl);
    });

    // Attach listeners
    listEl.querySelectorAll('.btn-study-deck').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt(e.target.getAttribute('data-index'), 10);
        launchSession([loadedDecks[index]], false);
      });
    });

    listEl.querySelectorAll('.btn-drill-deck').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt(e.target.getAttribute('data-index'), 10);
        launchSession([loadedDecks[index]], true);
      });
    });

    listEl.querySelectorAll('.btn-grid-deck').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt(e.target.getAttribute('data-index'), 10);
        openClassGrid(loadedDecks[index]);
      });
    });
  }

  function showImportConfirmation(importData) {
    currentImportPending = importData;

    document.getElementById('import-deck-name').value = importData.deckName;
    document.getElementById('import-stats-summary').textContent = `Found ${importData.cards.length} students in export.`;

    const warningsEl = document.getElementById('import-warnings');
    if (importData.warnings && importData.warnings.length > 0) {
      warningsEl.classList.remove('hidden');
      warningsEl.innerHTML = importData.warnings.map(w => `<div>⚠️ ${escapeHtml(w)}</div>`).join('');
    } else {
      warningsEl.classList.add('hidden');
      warningsEl.innerHTML = '';
    }

    const gridEl = document.getElementById('import-grid');
    gridEl.innerHTML = '';

    importData.cards.forEach((card, idx) => {
      const cardEl = document.createElement('div');
      cardEl.className = 'import-card-item' + (card.warning ? ' warn' : '');
      cardEl.innerHTML = `
        <img src="${card.photo}" class="import-photo-thumb" alt="Thumb">
        <div class="import-card-name">${escapeHtml(card.name)}</div>
        <div class="import-card-pref-wrapper">
          <input type="text" class="input-text sm import-preferred-input" data-idx="${idx}" placeholder="Preferred Name" value="${escapeHtml(card.preferredName || '')}">
        </div>
        ${card.warning ? `<div class="import-card-status">${escapeHtml(card.warning)}</div>` : ''}
      `;
      gridEl.appendChild(cardEl);
    });

    showScreen('import');
  }

  function launchSession(decks, isDrillEverything) {
    const session = startStudySession(decks, isDrillEverything);
    if (session.queue.length === 0) {
      alert('No cards are currently due in this session! Use "Drill All" to review anyway.');
      renderHome();
      showScreen('home');
      return;
    }
    renderStudyCard();
    showScreen('study');
  }

  function updateStudentNameDatalist(decks) {
    const datalist = document.getElementById('deck-student-names');
    if (!datalist) return;
    const names = new Set();

    if (activeSession && activeSession.queue && activeSession.currentIndex < activeSession.queue.length) {
      // Filter out students who have already been dialed in (answered) in this session
      const remainingItems = activeSession.queue.slice(activeSession.currentIndex);
      remainingItems.forEach(item => {
        const c = item.card;
        if (c.name && c.name !== 'Unknown Student') {
          names.add(c.name);
        }
        if (c.preferredName) {
          names.add(c.preferredName);
        }
      });
    } else if (decks) {
      decks.forEach(d => {
        if (d.cards) {
          d.cards.forEach(c => {
            if (c.name && c.name !== 'Unknown Student') {
              names.add(c.name);
            }
            if (c.preferredName) {
              names.add(c.preferredName);
            }
          });
        }
      });
    }

    datalist.innerHTML = Array.from(names)
      .sort()
      .map(n => `<option value="${escapeHtml(n)}"></option>`)
      .join('');
  }

  function renderStudyCard() {
    if (!activeSession || activeSession.currentIndex >= activeSession.queue.length) {
      renderSummary();
      showScreen('summary');
      return;
    }

    const currentItem = activeSession.queue[activeSession.currentIndex];
    const card = currentItem.card;

    updateStudentNameDatalist(activeSession.decks);

    document.getElementById('study-deck-title').textContent = currentItem.deck.deckName;
    document.getElementById('study-count-correct').textContent = activeSession.stats.correct;
    document.getElementById('study-count-partial').textContent = activeSession.stats.partial;
    document.getElementById('study-count-missed').textContent = activeSession.stats.missed;

    const progressPct = (activeSession.currentIndex / activeSession.queue.length) * 100;
    document.getElementById('study-progress-bar').style.width = progressPct + '%';

    document.getElementById('study-photo').src = card.photo;
    document.getElementById('study-feedback-area').innerHTML = '';

    const inputEl = document.getElementById('study-input');
    inputEl.value = '';
    inputEl.disabled = false;
    inputEl.focus();

    const submitBtn = document.getElementById('btn-submit-answer');
    submitBtn.textContent = 'Submit';
    submitBtn.setAttribute('data-state', 'submit');
  }

  function handleStudySubmit(e) {
    if (e) e.preventDefault();
    const submitBtn = document.getElementById('btn-submit-answer');
    const state = submitBtn.getAttribute('data-state');
    const inputEl = document.getElementById('study-input');

    if (state === 'advance') {
      activeSession.currentIndex++;
      renderStudyCard();
      return;
    }

    // Submit answer
    const typedText = inputEl.value;
    const currentItem = activeSession.queue[activeSession.currentIndex];
    const grade = gradeAnswer(typedText, currentItem.card);

    recordCardResult(grade.outcome);

    // Show feedback
    const feedbackEl = document.getElementById('study-feedback-area');
    let badgeHtml = '';
    if (grade.outcome === 'correct') {
      badgeHtml = `<span class="feedback-badge correct">✓ Correct</span> <span class="feedback-target">${escapeHtml(grade.targetSpelling)}</span>`;
    } else if (grade.outcome === 'partial') {
      badgeHtml = `<span class="feedback-badge partial">~ Partial</span> Full Name: <span class="feedback-target">${escapeHtml(grade.targetSpelling)}</span>`;
    } else {
      badgeHtml = `<span class="feedback-badge missed">✗ Missed</span> Correct Answer: <span class="feedback-target">${escapeHtml(grade.targetSpelling)}</span>`;
    }
    feedbackEl.innerHTML = badgeHtml;

    inputEl.disabled = true;
    submitBtn.textContent = 'Next Card (Space / Enter)';
    submitBtn.setAttribute('data-state', 'advance');
  }

  function handleRevealAnswer() {
    const submitBtn = document.getElementById('btn-submit-answer');
    if (submitBtn.getAttribute('data-state') === 'advance') {
      activeSession.currentIndex++;
      renderStudyCard();
      return;
    }

    const currentItem = activeSession.queue[activeSession.currentIndex];
    recordCardResult('miss');

    const feedbackEl = document.getElementById('study-feedback-area');
    feedbackEl.innerHTML = `<span class="feedback-badge missed">✗ Missed</span> Correct Answer: <span class="feedback-target">${escapeHtml(currentItem.card.name)}</span>`;

    const inputEl = document.getElementById('study-input');
    inputEl.disabled = true;
    submitBtn.textContent = 'Next Card (Space / Enter)';
    submitBtn.setAttribute('data-state', 'advance');
  }

  // --- THEME MANAGEMENT ---
  let currentTheme = 'dark';
  try {
    const savedTheme = localStorage.getItem('name_learner_theme');
    if (savedTheme === 'light' || savedTheme === 'dark') currentTheme = savedTheme;
  } catch (e) {}

  function applyTheme(theme) {
    currentTheme = theme;
    if (theme === 'light') {
      document.body.classList.add('theme-light');
      const btn = document.getElementById('btn-theme-toggle');
      if (btn) btn.innerHTML = '🌙 Dark Mode';
    } else {
      document.body.classList.remove('theme-light');
      const btn = document.getElementById('btn-theme-toggle');
      if (btn) btn.innerHTML = '☀️ Light Mode';
    }
    try {
      localStorage.setItem('name_learner_theme', theme);
    } catch (e) {}
  }

  function toggleTheme() {
    applyTheme(currentTheme === 'light' ? 'dark' : 'light');
  }

  // --- AUDIO SYNTHESIZER (Web Audio API) ---
  let soundMuted = false;
  let audioCtx = null;

  function getAudioContext() {
    if (!audioCtx) {
      const AudioContextClass = window.AudioContext || window.webkitAudioContext;
      if (AudioContextClass) audioCtx = new AudioContextClass();
    }
    if (audioCtx && audioCtx.state === 'suspended') {
      audioCtx.resume();
    }
    return audioCtx;
  }

  function playTone(freq, type, duration, delay = 0, startGain = 0.2, endGain = 0.001) {
    if (soundMuted) return;
    try {
      const ctx = getAudioContext();
      if (!ctx) return;
      setTimeout(() => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = type;
        osc.frequency.setValueAtTime(freq, ctx.currentTime);
        gain.gain.setValueAtTime(startGain, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(endGain, ctx.currentTime + duration);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + duration);
      }, delay);
    } catch (e) {}
  }

  function playSlideTone(startFreq, endFreq, type, duration, delay = 0) {
    if (soundMuted) return;
    try {
      const ctx = getAudioContext();
      if (!ctx) return;
      setTimeout(() => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = type;
        osc.frequency.setValueAtTime(startFreq, ctx.currentTime);
        osc.frequency.exponentialRampToValueAtTime(endFreq, ctx.currentTime + duration);
        gain.gain.setValueAtTime(0.25, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + duration);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + duration);
      }, delay);
    } catch (e) {}
  }

  function playTierSound(tierKey) {
    if (soundMuted) return;
    if (tierKey === 'distinction') {
      playTone(523.25, 'triangle', 0.15, 0);     // C5
      playTone(659.25, 'triangle', 0.15, 120);   // E5
      playTone(783.99, 'triangle', 0.15, 240);   // G5
      playTone(1046.50, 'triangle', 0.45, 360);  // C6
    } else if (tierKey === 'meeting') {
      playTone(392.00, 'sine', 0.15, 0);
      playTone(523.25, 'sine', 0.15, 120);
      playTone(659.25, 'sine', 0.35, 240);
    } else if (tierKey === 'developing') {
      playTone(440.00, 'sine', 0.15, 0);
      playTone(554.37, 'sine', 0.25, 140);
    } else if (tierKey === 'emerging') {
      playTone(349.23, 'sine', 0.15, 0);
      playTone(440.00, 'sine', 0.2, 120);
    } else if (tierKey === 'not-evident') {
      playSlideTone(260, 120, 'sawtooth', 0.4, 0);
    }
  }

  // --- CONFETTI ANIMATION ---
  let confettiAnimationId = null;
  function triggerConfetti() {
    const canvas = document.getElementById('summary-confetti-canvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;

    const colors = ['#eab308', '#22c55e', '#3b82f6', '#ec4899', '#a855f7', '#f97316'];
    const particles = [];
    for (let i = 0; i < 90; i++) {
      particles.push({
        x: Math.random() * canvas.width,
        y: Math.random() * canvas.height * 0.3 - canvas.height * 0.2,
        r: Math.random() * 6 + 4,
        d: Math.random() * 85,
        color: colors[Math.floor(Math.random() * colors.length)],
        tilt: Math.floor(Math.random() * 10) - 10,
        tiltAngleIncremental: Math.random() * 0.07 + 0.04,
        tiltAngle: 0
      });
    }

    let frames = 0;
    if (confettiAnimationId) cancelAnimationFrame(confettiAnimationId);

    function draw() {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      particles.forEach((p) => {
        p.tiltAngle += p.tiltAngleIncremental;
        p.y += (Math.cos(p.d) + 3 + p.r / 2) / 2;
        p.x += Math.sin(p.d);
        p.tilt = Math.sin(p.tiltAngle) * 15;

        ctx.beginPath();
        ctx.lineWidth = p.r;
        ctx.strokeStyle = p.color;
        ctx.moveTo(p.x + p.tilt + p.r / 2, p.y);
        ctx.lineTo(p.x + p.tilt, p.y + p.tilt + p.r / 2);
        ctx.stroke();
      });

      frames++;
      if (frames < 220) {
        confettiAnimationId = requestAnimationFrame(draw);
      } else {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
    }
    draw();
  }

  function renderSummary() {
    document.getElementById('sum-stat-correct').textContent = activeSession.stats.correct;
    document.getElementById('sum-stat-partial').textContent = activeSession.stats.partial;
    document.getElementById('sum-stat-missed').textContent = activeSession.stats.missed;

    // Record last run timestamp, score, and stats for each deck studied in this session
    const nowIso = new Date().toISOString();
    if (activeSession && activeSession.deckStatsMap && activeSession.deckStatsMap.size > 0) {
      activeSession.deckStatsMap.forEach((dStats, deck) => {
        if (dStats.total > 0) {
          const deckScorePct = Math.round(((dStats.correct + 0.5 * dStats.partial) / dStats.total) * 100);
          deck.lastStudiedAt = nowIso;
          deck.lastScore = deckScorePct;
          deck.lastRunStats = {
            correct: dStats.correct,
            partial: dStats.partial,
            missed: dStats.missed,
            total: dStats.total
          };
          deck.updatedAt = nowIso;
        }
      });
      setUnsaved(true);
      persistDecksToStorage();
    } else if (activeSession && activeSession.decks) {
      activeSession.decks.forEach(deck => {
        deck.lastStudiedAt = nowIso;
        deck.updatedAt = nowIso;
      });
      setUnsaved(true);
      persistDecksToStorage();
    }

    // Performance rubric tier calculation
    const totalSessionCards = activeSession.stats.correct + activeSession.stats.partial + activeSession.stats.missed;
    const score = totalSessionCards > 0 ? ((activeSession.stats.correct + 0.5 * activeSession.stats.partial) / totalSessionCards) * 100 : 0;

    let tierKey = 'not-evident';
    let tierTitle = 'Not Yet Evident';
    let tierAvatar = '🫥';
    let animClass = 'anim-wobble-shake';
    let tierDesc = 'Keep at it! Learning new faces and names takes repetition.';

    if (score >= 95) {
      tierKey = 'distinction';
      tierTitle = 'Meeting with Distinction';
      tierAvatar = '👑';
      animClass = 'anim-crown-triumph';
      tierDesc = 'Flawless performance! You have mastered these student names!';
    } else if (score >= 80) {
      tierKey = 'meeting';
      tierTitle = 'Meeting';
      tierAvatar = '⭐';
      animClass = 'anim-pop-star';
      tierDesc = 'Great job! You have reached target mastery for this class!';
    } else if (score >= 55) {
      tierKey = 'developing';
      tierTitle = 'Developing';
      tierAvatar = '🚀';
      animClass = 'anim-float-sparkle';
      tierDesc = 'Solid progress! The connections are building nicely.';
    } else if (score >= 30) {
      tierKey = 'emerging';
      tierTitle = 'Emerging';
      tierAvatar = '🌱';
      animClass = 'anim-bounce-gently';
      tierDesc = 'You are on your way! A couple more sessions will lock these in.';
    }

    const bannerEl = document.getElementById('summary-performance-banner');
    if (bannerEl) {
      bannerEl.className = `summary-performance-banner tier-${tierKey}`;
      bannerEl.innerHTML = `
        <div class="tier-avatar ${animClass}">${tierAvatar}</div>
        <div>
          <div class="tier-title-badge">${escapeHtml(tierTitle)}</div>
          <div class="tier-description">${escapeHtml(tierDesc)}</div>
        </div>
      `;
    }

    playTierSound(tierKey);
    if (tierKey === 'distinction' || tierKey === 'meeting') {
      triggerConfetti();
    }

    // Box distribution across all cards in active decks
    const boxCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
    let totalCards = 0;
    activeSession.decks.forEach(deck => {
      deck.cards.forEach(card => {
        const b = Math.min(Math.max(1, card.box || 1), 5);
        boxCounts[b]++;
        totalCards++;
      });
    });

    const distEl = document.getElementById('box-distribution');
    distEl.innerHTML = '';
    for (let b = 1; b <= 5; b++) {
      const cnt = boxCounts[b];
      const pct = totalCards > 0 ? (cnt / totalCards) * 100 : 0;
      const barWrapper = document.createElement('div');
      barWrapper.className = 'box-bar-wrapper';
      barWrapper.innerHTML = `
        <div class="box-bar" style="height: ${Math.max(10, pct)}%;"></div>
        <div>Box ${b}: ${cnt}</div>
      `;
      distEl.appendChild(barWrapper);
    }

    // Missed cards list with delete option (§8.4)
    const missedListEl = document.getElementById('missed-cards-list');
    const missedSection = document.getElementById('missed-cards-section');

    if (activeSession.missedCards.length === 0) {
      missedSection.classList.add('hidden');
    } else {
      missedSection.classList.remove('hidden');
      missedListEl.innerHTML = '';

      activeSession.missedCards.forEach((item) => {
        const rowEl = document.createElement('div');
        rowEl.className = 'missed-card-row';
        rowEl.innerHTML = `
          <div class="missed-card-info">
            <img src="${item.card.photo}" class="missed-card-thumb" alt="Thumb">
            <span class="missed-card-name">${escapeHtml(item.card.name)}</span>
          </div>
          <button class="btn danger btn-delete-card" data-card-id="${item.card.id}">Delete Card</button>
        `;
        missedListEl.appendChild(rowEl);
      });

      missedListEl.querySelectorAll('.btn-delete-card').forEach(btn => {
        btn.addEventListener('click', (e) => {
          const cardId = e.target.getAttribute('data-card-id');
          // Delete card from deck
          activeSession.decks.forEach(deck => {
            deck.cards = deck.cards.filter(c => c.id !== cardId);
          });
          e.target.closest('.missed-card-row').remove();
          setUnsaved(true);
          persistDecksToStorage();
        });
      });
    }
  }

  // --- CLASS ROSTER GRID VIEW ---
  function openClassGrid(deck) {
    activeGridDeck = deck;
    gridNamesHidden = true;
    gridRevealedIds.clear();

    document.getElementById('grid-deck-title').textContent = `${deck.deckName} - Class Roster Grid (${deck.cards.length} Students)`;
    document.getElementById('btn-toggle-all-names').textContent = 'Show All Names';

    const toggleStudyBtn = document.getElementById('btn-grid-toggle-study');
    if (toggleStudyBtn) {
      toggleStudyBtn.classList.remove('hidden');
      if (activeSession && activeSession.queue && activeSession.currentIndex < activeSession.queue.length) {
        toggleStudyBtn.textContent = 'Resume Study';
      } else {
        toggleStudyBtn.textContent = 'Study Deck';
      }
    }

    renderClassGrid();
    showScreen('grid');
  }

  function renderClassGrid() {
    if (!activeGridDeck) return;
    const container = document.getElementById('class-grid-container');
    container.innerHTML = '';

    if (!activeGridDeck.cards || activeGridDeck.cards.length === 0) {
      container.innerHTML = '<div class="empty-state">No students found in this deck.</div>';
      return;
    }

    activeGridDeck.cards.forEach(card => {
      const isHidden = gridNamesHidden && !gridRevealedIds.has(card.id);
      const isEditing = editingCardId === card.id;

      const cardEl = document.createElement('div');
      cardEl.className = 'grid-student-card' + (isEditing ? ' editing' : '');

      if (isEditing) {
        cardEl.innerHTML = `
          <img src="${card.photo}" class="grid-student-photo" alt="Photo of ${escapeHtml(card.name)}">
          <div class="grid-student-edit-box">
            <div class="grid-edit-roster-name" title="${escapeHtml(card.name)}">${escapeHtml(card.name)}</div>
            <input type="text" class="input-text sm grid-pref-input" placeholder="Preferred Name" value="${escapeHtml(card.preferredName || '')}">
            <div class="grid-edit-btn-row">
              <button type="button" class="btn primary sm btn-save-pref" title="Save preferred name">✓ Save</button>
              <button type="button" class="btn secondary sm btn-cancel-pref" title="Cancel">✕</button>
            </div>
          </div>
          <div class="grid-box-tag">Box ${card.box || 1}</div>
        `;

        const prefInput = cardEl.querySelector('.grid-pref-input');
        const saveBtn = cardEl.querySelector('.btn-save-pref');
        const cancelBtn = cardEl.querySelector('.btn-cancel-pref');

        cardEl.querySelectorAll('input, button, .grid-student-edit-box').forEach(el => {
          el.addEventListener('click', (e) => e.stopPropagation());
        });

        prefInput.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            saveBtn.click();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            cancelBtn.click();
          }
        });

        saveBtn.addEventListener('click', () => {
          const val = prefInput.value.trim();
          card.preferredName = val || null;
          activeGridDeck.updatedAt = new Date().toISOString();
          setUnsaved(true);
          updateStudentNameDatalist(loadedDecks);
          editingCardId = null;
          renderClassGrid();
        });

        cancelBtn.addEventListener('click', () => {
          editingCardId = null;
          renderClassGrid();
        });

        setTimeout(() => prefInput.focus(), 50);
      } else {
        const prefText = card.preferredName ? `<div class="grid-student-pref-name">"${escapeHtml(card.preferredName)}"</div>` : '';
        cardEl.innerHTML = `
          <img src="${card.photo}" class="grid-student-photo" alt="Photo of ${escapeHtml(card.name)}">
          <div class="grid-student-name-wrapper ${isHidden ? 'hidden-name' : ''}">
            <div class="grid-student-name">${escapeHtml(card.name)}</div>
            ${prefText}
            <button type="button" class="grid-edit-pencil" title="Edit preferred name" data-id="${card.id}">✏️</button>
          </div>
          <div class="grid-box-tag">Box ${card.box || 1}</div>
        `;

        const editBtn = cardEl.querySelector('.grid-edit-pencil');
        if (editBtn) {
          editBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            editingCardId = card.id;
            renderClassGrid();
          });
        }

        cardEl.addEventListener('click', () => {
          if (gridRevealedIds.has(card.id)) {
            gridRevealedIds.delete(card.id);
          } else {
            gridRevealedIds.add(card.id);
          }
          renderClassGrid();
        });
      }

      container.appendChild(cardEl);
    });
  }
  async function exportCollectiveJSON(decksToExport, suggestedName) {
    const targetDecks = decksToExport && decksToExport.length > 0 ? decksToExport : loadedDecks;
    if (!targetDecks || targetDecks.length === 0) return;

    const exportData = {
      schema: 1,
      isMultiDeck: true,
      exportedAt: new Date().toISOString(),
      decks: targetDecks.map(deck => ({
        schema: 1,
        deckName: deck.deckName,
        createdAt: deck.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        sessionCount: deck.sessionCount || 0,
        lastStudiedAt: deck.lastStudiedAt || null,
        lastScore: deck.lastScore !== undefined ? deck.lastScore : null,
        lastRunStats: deck.lastRunStats || null,
        cards: deck.cards.map(c => ({
          id: c.id,
          name: c.name,
          preferredName: c.preferredName || null,
          photo: c.photo,
          box: c.box || 1,
          dueSession: c.dueSession || 1,
          seen: c.seen || 0,
          correct: c.correct || 0,
          missed: c.missed || 0
        }))
      }))
    };

    const jsonStr = JSON.stringify(exportData, null, 2);
    const dateStr = new Date().toISOString().slice(0, 10);
    const defaultFileName = suggestedName || (targetDecks.length === 1 
      ? `${targetDecks[0].deckName.replace(/[^\w\s-]/g, '')}.deck.json` 
      : `All_Classes_Collective_${dateStr}.deck.json`);

    let savedViaAPI = false;
    if ('showSaveFilePicker' in window) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: defaultFileName,
          types: [{ description: 'Collective Deck JSON File', accept: { 'application/json': ['.json'] } }]
        });
        const writable = await handle.createWritable();
        await writable.write(jsonStr);
        await writable.close();
        savedViaAPI = true;
      } catch (err) {
        if (err.name !== 'AbortError') {
          console.warn('showSaveFilePicker failed or cancelled, falling back to download link:', err);
        } else {
          return;
        }
      }
    }

    if (!savedViaAPI) {
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = defaultFileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    setUnsaved(false);
  }

  async function saveDecksToDisk() {
    if (loadedDecks.length === 0) return;

    for (const deck of loadedDecks) {
      const cleanDeck = {
        schema: 1,
        deckName: deck.deckName,
        createdAt: deck.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        sessionCount: deck.sessionCount || 0,
        lastStudiedAt: deck.lastStudiedAt || null,
        lastScore: deck.lastScore !== undefined ? deck.lastScore : null,
        lastRunStats: deck.lastRunStats || null,
        cards: deck.cards.map(c => ({
          id: c.id,
          name: c.name,
          preferredName: c.preferredName || null,
          photo: c.photo,
          box: c.box || 1,
          dueSession: c.dueSession || 1,
          seen: c.seen || 0,
          correct: c.correct || 0,
          missed: c.missed || 0
        }))
      };

      const jsonStr = JSON.stringify(cleanDeck, null, 2);
      const fileName = `${deck.deckName.replace(/[^\w\s-]/g, '')}.deck.json`;

      let savedViaAPI = false;
      if ('showSaveFilePicker' in window) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: fileName,
            types: [{ description: 'Deck JSON File', accept: { 'application/json': ['.json'] } }]
          });
          const writable = await handle.createWritable();
          await writable.write(jsonStr);
          await writable.close();
          savedViaAPI = true;
        } catch (err) {
          if (err.name !== 'AbortError') {
            console.warn('showSaveFilePicker failed or cancelled, falling back to download link:', err);
          } else {
            continue; // User cancelled save dialog
          }
        }
      }

      if (!savedViaAPI) {
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    }

    setUnsaved(false);
  }

  let pdfImportQueue = [];

  async function handleBatchFileImport(files) {
    if (!files || files.length === 0) return;

    let jsonCount = 0;
    const importedDeckNames = [];
    const pdfFiles = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const fileName = file.name;
      if (fileName.toLowerCase().endsWith('.deck.json') || fileName.toLowerCase().endsWith('.json')) {
        try {
          const text = await file.text();
          const data = JSON.parse(text);
          let decksToImport = [];

          if (Array.isArray(data)) {
            decksToImport = data.filter(d => d && d.cards && Array.isArray(d.cards));
          } else if (data && data.decks && Array.isArray(data.decks)) {
            decksToImport = data.decks.filter(d => d && d.cards && Array.isArray(d.cards));
          } else if (data && data.cards && Array.isArray(data.cards)) {
            decksToImport = [data];
          }

          if (decksToImport.length > 0) {
            decksToImport.forEach(deckObj => {
              const existingIdx = loadedDecks.findIndex(d => d.deckName === deckObj.deckName);
              if (existingIdx >= 0) {
                loadedDecks[existingIdx] = deckObj;
              } else {
                loadedDecks.push(deckObj);
              }
              if (deckObj.deckName) {
                importedDeckNames.push(deckObj.deckName);
              }
              jsonCount++;
            });
          } else {
            alert(`No valid card decks found in ${fileName}.`);
          }
        } catch (err) {
          alert(`Failed to load ${fileName}: ` + err.message);
        }
      } else if (fileName.toLowerCase().endsWith('.pdf')) {
        pdfFiles.push(file);
      } else {
        alert(`Unsupported file format (${fileName}). Please provide .pdf or .deck.json files.`);
      }
    }

    if (jsonCount > 0) {
      renderHome();
      const deckListStr = importedDeckNames.length > 0 ? `: ${importedDeckNames.join(', ')}` : '.';
      alert(`Successfully loaded ${jsonCount} deck(s) from JSON${deckListStr}`);
    }

    // Process PDF files
    for (const pdfFile of pdfFiles) {
      try {
        const arrayBuffer = await pdfFile.arrayBuffer();
        const importResult = await parsePDF(arrayBuffer, pdfFile.name);

        const existingDeck = loadedDecks.find(d => d.deckName.toLowerCase() === importResult.deckName.toLowerCase());
        if (existingDeck) {
          if (confirm(`A deck named "${importResult.deckName}" is already loaded. Would you like to MERGE the new PDF into it?\n\nMerging preserves existing progress for returning students.`)) {
            mergeDecks(existingDeck, importResult.cards);
            setUnsaved(true);
            persistDecksToStorage();
            renderHome();
            continue;
          }
        }
        pdfImportQueue.push(importResult);
      } catch (err) {
        alert(`Failed to parse PDF (${pdfFile.name}): ` + err.message);
      }
    }

    if (pdfImportQueue.length > 0 && !currentImportPending) {
      processNextImportQueue();
    }
  }

  function processNextImportQueue() {
    if (pdfImportQueue.length === 0) {
      currentImportPending = null;
      renderHome();
      showScreen('home');
      return;
    }
    const nextImport = pdfImportQueue.shift();
    showImportConfirmation(nextImport);
  }

  function mergeDecks(targetDeck, newCards) {
    const existingCardsMap = new Map();
    targetDeck.cards.forEach(c => existingCardsMap.set(c.name, c));

    const mergedCards = [];
    newCards.forEach(nCard => {
      if (existingCardsMap.has(nCard.name)) {
        mergedCards.push(existingCardsMap.get(nCard.name));
        existingCardsMap.delete(nCard.name);
      } else {
        mergedCards.push(nCard);
      }
    });

    targetDeck.cards = mergedCards;
    targetDeck.updatedAt = new Date().toISOString();
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // --- INITIALIZATION & BINDINGS ---
  function init() {
    initPdfWorker();
    applyTheme(currentTheme);

    // Theme & Sound Toggles
    const themeToggleBtn = document.getElementById('btn-theme-toggle');
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener('click', toggleTheme);
    }

    const soundToggleBtn = document.getElementById('btn-sound-toggle');
    if (soundToggleBtn) {
      soundToggleBtn.addEventListener('click', () => {
        soundMuted = !soundMuted;
        soundToggleBtn.textContent = soundMuted ? '🔇 Sound Off' : '🔊 Sound On';
      });
    }

    // File Drop Zone
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const btnBrowse = document.getElementById('btn-browse');

    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
      }, false);
    });

    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt.files && dt.files.length > 0) {
        handleBatchFileImport(dt.files);
      }
    });

    btnBrowse.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleBatchFileImport(e.target.files);
      }
    });

    // Study Combined
    document.getElementById('btn-study-combined').addEventListener('click', () => {
      const selectedBoxes = document.querySelectorAll('.deck-checkbox:checked');
      if (selectedBoxes.length === 0) {
        alert('Please select two or more decks to study combined.');
        return;
      }
      const selectedDecks = Array.from(selectedBoxes).map(cb => loadedDecks[parseInt(cb.getAttribute('data-index'), 10)]);
      launchSession(selectedDecks, false);
    });

    // Import confirmation actions
    document.getElementById('btn-confirm-import').addEventListener('click', () => {
      if (!currentImportPending) return;
      const finalName = document.getElementById('import-deck-name').value.trim() || currentImportPending.deckName;

      // Save preferred names from import grid inputs
      const prefInputs = document.querySelectorAll('.import-preferred-input');
      prefInputs.forEach(input => {
        const idx = parseInt(input.getAttribute('data-idx'), 10);
        if (!isNaN(idx) && currentImportPending.cards[idx]) {
          const val = input.value.trim();
          currentImportPending.cards[idx].preferredName = val || null;
        }
      });

      const newDeck = {
        schema: 1,
        deckName: finalName,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        sessionCount: 0,
        lastStudiedAt: null,
        lastScore: null,
        lastRunStats: null,
        cards: currentImportPending.cards
      };

      const existingIdx = loadedDecks.findIndex(d => d.deckName === finalName);
      if (existingIdx >= 0) {
        loadedDecks[existingIdx] = newDeck;
      } else {
        loadedDecks.push(newDeck);
      }

      setUnsaved(true);
      persistDecksToStorage();
      currentImportPending = null;
      processNextImportQueue();
    });

    document.getElementById('btn-cancel-import').addEventListener('click', () => {
      currentImportPending = null;
      processNextImportQueue();
    });

    // Study controls
    document.getElementById('study-form').addEventListener('submit', handleStudySubmit);
    document.getElementById('btn-reveal-answer').addEventListener('click', handleRevealAnswer);
    document.getElementById('btn-exit-study').addEventListener('click', () => {
      renderHome();
      showScreen('home');
    });

    const studyViewGridBtn = document.getElementById('btn-study-view-grid');
    if (studyViewGridBtn) {
      studyViewGridBtn.addEventListener('click', () => {
        if (activeSession && activeSession.queue && activeSession.queue.length > 0) {
          const currentItem = activeSession.queue[activeSession.currentIndex] || activeSession.queue[0];
          openClassGrid(currentItem.deck);
        }
      });
    }

    const photoContainer = document.getElementById('study-photo-container');
    if (photoContainer) {
      photoContainer.addEventListener('click', () => {
        const submitBtn = document.getElementById('btn-submit-answer');
        const state = submitBtn ? submitBtn.getAttribute('data-state') : null;
        if (state === 'advance') {
          activeSession.currentIndex++;
          renderStudyCard();
        } else {
          handleRevealAnswer();
        }
      });
    }

    // Class Grid controls
    document.getElementById('btn-exit-grid').addEventListener('click', () => {
      renderHome();
      showScreen('home');
    });

    const gridToggleStudyBtn = document.getElementById('btn-grid-toggle-study');
    if (gridToggleStudyBtn) {
      gridToggleStudyBtn.addEventListener('click', () => {
        if (activeSession && activeSession.queue && activeSession.currentIndex < activeSession.queue.length) {
          showScreen('study');
        } else if (activeGridDeck) {
          launchSession([activeGridDeck], false);
        }
      });
    }

    document.getElementById('btn-toggle-all-names').addEventListener('click', () => {
      gridNamesHidden = !gridNamesHidden;
      gridRevealedIds.clear();
      const btn = document.getElementById('btn-toggle-all-names');
      btn.textContent = gridNamesHidden ? 'Show All Names' : 'Hide All Names';
      renderClassGrid();
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      if (document.body.className === 'screen-study') {
        const submitBtn = document.getElementById('btn-submit-answer');
        const state = submitBtn ? submitBtn.getAttribute('data-state') : null;
        const inputEl = document.getElementById('study-input');

        // Enter key shortcut fix
        if (e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter') {
          if (state === 'advance') {
            e.preventDefault();
            if (document.activeElement) document.activeElement.blur();
            activeSession.currentIndex++;
            renderStudyCard();
            return;
          }
        }

        // Space key shortcut fix
        const isSpace = e.key === ' ' || e.code === 'Space' || e.key === 'Spacebar';
        if (isSpace) {
          if (state === 'advance') {
            e.preventDefault();
            if (document.activeElement) document.activeElement.blur();
            activeSession.currentIndex++;
            renderStudyCard();
            return;
          } else if (document.activeElement === inputEl && inputEl.value === '') {
            e.preventDefault();
            if (document.activeElement) document.activeElement.blur();
            handleRevealAnswer();
            return;
          } else if (document.activeElement !== inputEl) {
            e.preventDefault();
            if (document.activeElement) document.activeElement.blur();
            if (state === 'advance') {
              activeSession.currentIndex++;
              renderStudyCard();
            } else {
              handleRevealAnswer();
            }
            return;
          }
        }

        // Escape key shortcut
        if (e.code === 'Escape') {
          renderHome();
          showScreen('home');
        }
      } else if (document.body.className === 'screen-grid') {
        if (e.code === 'Escape') {
          renderHome();
          showScreen('home');
        }
      }
    });

    // Summary screen actions
    document.getElementById('btn-save-decks').addEventListener('click', saveDecksToDisk);

    const collectiveSummaryBtn = document.getElementById('btn-save-collective');
    if (collectiveSummaryBtn) {
      collectiveSummaryBtn.addEventListener('click', () => {
        const decksToExport = activeSession && activeSession.decks ? activeSession.decks : loadedDecks;
        exportCollectiveJSON(decksToExport);
      });
    }

    const collectiveHomeBtn = document.getElementById('btn-export-collective-home');
    if (collectiveHomeBtn) {
      collectiveHomeBtn.addEventListener('click', () => {
        exportCollectiveJSON(loadedDecks);
      });
    }

    const restartSessionBtn = document.getElementById('btn-restart-session');
    if (restartSessionBtn) {
      restartSessionBtn.addEventListener('click', () => {
        if (activeSession && activeSession.decks) {
          launchSession(activeSession.decks, true);
        }
      });
    }

    document.getElementById('btn-continue-studying').addEventListener('click', () => {
      renderHome();
      showScreen('home');
    });

    // Warn on close if unsaved (§8.5)
    window.addEventListener('beforeunload', (e) => {
      if (hasUnsavedChanges) {
        e.preventDefault();
        e.returnValue = '';
      }
    });

    loadDecksFromStorage().then(stored => {
      if (stored && Array.isArray(stored) && stored.length > 0) {
        loadedDecks = stored;
      }
      renderHome();
    }).catch(err => {
      console.warn('Storage load failed:', err);
      renderHome();
    });
  }

  window.addEventListener('DOMContentLoaded', init);
})();
