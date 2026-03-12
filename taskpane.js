/* global Word, Office */

// Language mappings
const LANGUAGE_MAP = {
    'auto': 'Auto-detect',
    'en': 'English',
    'ar': 'Arabic',
    'es': 'Spanish',
    'fr': 'French',
    'de': 'German',
    'it': 'Italian',
    'pt': 'Portuguese',
    'ru': 'Russian',
    'ja': 'Japanese',
    'ko': 'Korean',
    'zh': 'Chinese (Simplified)',
    'zh-TW': 'Chinese (Traditional)',
    'hi': 'Hindi',
    'tr': 'Turkish',
    'nl': 'Dutch',
    'pl': 'Polish',
    'sv': 'Swedish',
    'da': 'Danish',
    'no': 'Norwegian',
    'fi': 'Finnish',
    'el': 'Greek',
    'he': 'Hebrew',
    'th': 'Thai',
    'vi': 'Vietnamese',
    'id': 'Indonesian',
    'ms': 'Malay',
    'cs': 'Czech',
    'hu': 'Hungarian',
    'ro': 'Romanian',
    'uk': 'Ukrainian'
};

// Model pricing (per 1M tokens)
const MODEL_PRICING = {
    'gpt-4.1': { input: 3.00, output: 12.00 },      // Main office model
    'gpt-5.2': { input: 20.00, output: 80.00 },     // Premium for complex files
    'gpt-4o-mini': { input: 0.150, output: 0.600 }  // Fast & economical
};

// Build system prompt based on translation mode and style
function buildSystemPrompt(fromLang, toLang, targetLanguage, mode, style) {
    let basePrompt = '';
    let styleInstructions = '';
    
    // Special handling for casual mode - always natural regardless of style
    if (mode === 'casual') {
        basePrompt = `Translate conversationally like a native speaker in everyday language. Make it sound natural and friendly, as if texting or chatting. Use colloquial expressions when appropriate.`;
    } else {
        // Style instructions for non-casual modes
        if (style === 'strict') {
            styleInstructions = 'Translate literally and precisely. Maintain exact sentence structure. Use formal terminology.';
        } else if (style === 'human') {
            styleInstructions = 'Translate naturally as a native speaker would say it. Prioritize readability and natural flow over literal accuracy.';
        } else {
            styleInstructions = 'Balance accuracy with natural language. Keep it professional but readable.';
        }
        
        // Mode-specific prompts
        switch(mode) {
            case 'legal':
                basePrompt = `You are a certified legal translator. Use precise legal terminology. Preserve structure and formatting exactly. Keep all article numbers, dates, names, and IDs unchanged. ${styleInstructions}`;
                break;
            case 'certificate':
                basePrompt = `Translate in official government certificate style. Keep names, numbers, dates, seals, and stamps unchanged. Use formal government language. ${styleInstructions}`;
                break;
            case 'bank':
                basePrompt = `Translate using formal banking and financial terminology. Keep account numbers, amounts, dates, and reference codes unchanged. Use standard banking language. ${styleInstructions}`;
                break;
            case 'medical':
                basePrompt = `Translate medical reports using accurate medical terminology. Keep patient names, dates, test results, and measurements unchanged. Maintain clinical precision. ${styleInstructions}`;
                break;
            case 'academic':
                basePrompt = `Translate academic documents with scholarly terminology. Keep citations, dates, names, and numerical data unchanged. Maintain academic tone. ${styleInstructions}`;
                break;
            case 'business':
                basePrompt = `Translate business contracts and documents using formal business language. Keep company names, dates, amounts, and clause numbers unchanged. ${styleInstructions}`;
                break;
            case 'technical':
                basePrompt = `Translate technical manuals using precise technical terminology. Keep model numbers, specifications, measurements, and codes unchanged. ${styleInstructions}`;
                break;
            case 'government':
                basePrompt = `Translate official government documents using formal administrative language. Keep all reference numbers, dates, names, and official codes unchanged. ${styleInstructions}`;
                break;
            default:
                basePrompt = `You are a professional translator. ${styleInstructions}`;
        }
    }
    
    // Add auto-detect or source language
    if (fromLang === 'auto') {
        return `${basePrompt} Detect the source language and translate to ${targetLanguage}. 

CRITICAL FORMATTING RULES:
- Preserve the EXACT formatting, structure, and layout of the original text
- Do NOT add any numbering, bullets, or formatting that doesn't exist in the original
- If the original has numbering (1., 2., 3.), keep it exactly as is
- If the original has NO numbering, do NOT add any
- Maintain all line breaks, spacing, and indentation exactly as in the source
- IMPORTANT: Preserve all numbers, dates, IDs, and proper names exactly as they appear

Return ONLY the translated text without any explanations. At the very end, on a new line, write "DETECTED:" followed by the detected language name in English.`;
    } else {
        const sourceLanguage = LANGUAGE_MAP[fromLang];
        return `${basePrompt} Translate from ${sourceLanguage} to ${targetLanguage}. 

CRITICAL FORMATTING RULES:
- Preserve the EXACT formatting, structure, and layout of the original text
- Do NOT add any numbering, bullets, or formatting that doesn't exist in the original
- If the original has numbering (1., 2., 3.), keep it exactly as is
- If the original has NO numbering, do NOT add any
- Maintain all line breaks, spacing, and indentation exactly as in the source
- IMPORTANT: Preserve all numbers, dates, IDs, and proper names exactly as they appear

Return ONLY the translated text without any explanations.`;
    }
}

// Get appropriate temperature based on mode and style
function getTemperatureForMode(mode, style) {
    // Strict modes need lower temperature
    if (style === 'strict') return 0.1;
    if (style === 'human') return 0.4;
    
    // Mode-based defaults
    switch(mode) {
        case 'legal':
        case 'certificate':
        case 'bank':
        case 'medical':
        case 'government':
            return 0.1; // Very precise
        case 'academic':
        case 'business':
        case 'technical':
            return 0.2; // Precise but slightly flexible
        case 'casual':
            return 0.4; // More natural
        default:
            return 0.3; // Balanced
    }
}

// Initialize usage stats
let usageStats = {
    totalTranslations: 0,
    totalWords: 0,
    totalInputTokens: 0,
    totalOutputTokens: 0,
    totalCost: 0
};

// Sync all data to server (so dashboard always has latest)
async function syncToServer() {
    try {
        await fetch('http://localhost:3000/api/stats', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                usageStats,
                cacheStats,
                translationCache
            })
        });
    } catch(e) {
        // Server might not be running, silently ignore
    }
}

// Log a single translation to server
async function logToServer(logData) {
    try {
        await fetch('http://localhost:3000/api/log', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(logData)
        });
    } catch(e) {}
}

// Settings with defaults
let memorySettings = {
    useCache: true,
    replaceAll: false,
    lockNumbers: true
};

// In-memory cache object (synced with localStorage)
let translationCache = {};
let cacheStats = { hits: 0, savedCost: 0 };

function loadMemorySystem() {
    // Load settings
    const savedSettings = localStorage.getItem('memory_settings');
    if (savedSettings) {
        memorySettings = { ...memorySettings, ...JSON.parse(savedSettings) };
    }
    // Load cache
    const savedCache = localStorage.getItem('translation_cache');
    if (savedCache) {
        translationCache = JSON.parse(savedCache);
    }
    // Load cache stats
    const savedCacheStats = localStorage.getItem('cache_stats');
    if (savedCacheStats) {
        cacheStats = JSON.parse(savedCacheStats);
    }
    // Sync toggles UI
    syncToggleUI();
    updateCacheDisplay();
}

function syncToggleUI() {
    ['useCache', 'replaceAll', 'lockNumbers'].forEach(key => {
        const el = document.getElementById('toggle' + key.charAt(0).toUpperCase() + key.slice(1));
        if (el) {
            el.classList.toggle('on', !!memorySettings[key]);
        }
    });
}

function toggleSetting(key) {
    memorySettings[key] = !memorySettings[key];
    syncToggleUI();
    localStorage.setItem('memory_settings', JSON.stringify(memorySettings));
    showStatus(`${key === 'useCache' ? '🧠 Memory' : key === 'replaceAll' ? '🔄 Replace All' : '🔒 Lock Numbers'}: ${memorySettings[key] ? 'ON' : 'OFF'}`, 'info');
}

// Build cache key
function cacheKey(text, fromLang, toLang) {
    return `${fromLang}_${toLang}_${text.trim()}`;
}

// Lookup exact match in cache
function lookupCache(text, fromLang, toLang) {
    if (!memorySettings.useCache) return null;
    const key = cacheKey(text, fromLang, toLang);
    return translationCache[key] || null;
}

// Save to cache
function saveToCache(sourceText, translatedText, fromLang, toLang) {
    const key = cacheKey(sourceText, fromLang, toLang);
    translationCache[key] = translatedText;
    localStorage.setItem('translation_cache', JSON.stringify(translationCache));
    updateCacheDisplay();
    syncToServer();
}

// Save multiple pairs from chat mapping: { source: translation, ... }
function savePairsToCache(pairs, fromLang, toLang) {
    Object.entries(pairs).forEach(([src, tgt]) => {
        if (src.trim() && tgt.trim()) {
            translationCache[cacheKey(src, fromLang, toLang)] = tgt.trim();
        }
    });
    localStorage.setItem('translation_cache', JSON.stringify(translationCache));
    updateCacheDisplay();
    syncToServer();
}
function recordCacheHit(model) {
    cacheStats.hits++;
    // Estimate cost saved: ~50 tokens per phrase at gpt-4.1 pricing
    const pricing = MODEL_PRICING[model] || MODEL_PRICING['gpt-4.1'];
    cacheStats.savedCost += (50 / 1000000) * pricing.input + (50 / 1000000) * pricing.output;
    localStorage.setItem('cache_stats', JSON.stringify(cacheStats));
    updateCacheDisplay();
    syncToServer();
}
function clearCache() {
    if (!confirm('Clear all saved translation pairs?')) return;
    translationCache = {};
    cacheStats = { hits: 0, savedCost: 0 };
    localStorage.removeItem('translation_cache');
    localStorage.removeItem('cache_stats');
    updateCacheDisplay();
    showStatus('🗑️ Translation memory cleared', 'info');
    const preview = document.getElementById('cachePreview');
    if (preview) { preview.innerHTML = ''; preview.style.display = 'none'; }
}

// Update cache display in UI
function updateCacheDisplay() {
    const count = Object.keys(translationCache).length;
    const countEl = document.getElementById('cacheCount');
    const hitsEl = document.getElementById('cacheHits');
    const savedEl = document.getElementById('cacheSaved');
    if (countEl) countEl.textContent = count.toLocaleString();
    if (hitsEl) hitsEl.textContent = cacheStats.hits.toLocaleString();
    if (savedEl) savedEl.textContent = `$${(cacheStats.savedCost || 0).toFixed(4)}`;
}

// Toggle cache preview panel
function toggleCachePreview() {
    const preview = document.getElementById('cachePreview');
    if (preview.style.display === 'block') {
        preview.style.display = 'none';
        return;
    }
    const entries = Object.entries(translationCache);
    if (entries.length === 0) {
        preview.innerHTML = '<div style="color:#999; text-align:center; padding:8px;">No saved pairs yet</div>';
    } else {
        // Show last 20 entries
        const html = entries.slice(-20).map(([key, val]) => {
            const parts = key.split('_');
            const src = parts.slice(2).join('_'); // everything after lang codes
            return `<div class="cache-entry">${src}<span class="cache-arrow">→</span>${val}</div>`;
        }).join('');
        preview.innerHTML = html + (entries.length > 20 ? `<div style="color:#999; font-size:10px; text-align:center; padding:4px;">...and ${entries.length - 20} more</div>` : '');
    }
    preview.style.display = 'block';
}

// ===== APPLY CHAT → WORD =====
// Parse the last assistant chat message for source|translation pairs
// then replace the selected text (or all occurrences) in Word.
async function applyChatToWord() {
    const apiKey = localStorage.getItem('openai_api_key');

    // Find last assistant message content
    const lastAssistant = [...chatHistory].reverse().find(m => m.role === 'assistant');
    if (!lastAssistant) {
        showStatus('❌ No chat response to apply. Ask the assistant to translate first.', 'error');
        return;
    }

    // Parse pipe-separated pairs: "source | translation"
    const pairs = parseChatPairs(lastAssistant.content);
    const hasPairs = Object.keys(pairs).length > 0;

    // If no pairs found → treat entire chat response as direct translation for selected text
    if (!hasPairs) {
        try {
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.load("text");
                await context.sync();
                if (!range.text.trim()) { showStatus("⚠️ No text selected in Word", "error"); return; }
                range.insertText(lastAssistant.content.trim(), Word.InsertLocation.replace);
                await context.sync();
                showStatus("✅ Applied chat response to selection", "success");
            });
        } catch(err) { showStatus("❌ Error: " + err.message, "error"); }
        return;
    }

    const sourceLang = document.getElementById('sourceLang').value;
    const targetLang = document.getElementById('targetLang').value;
    const model = document.getElementById('modelSelect').value;

    // Save to cache first
    savePairsToCache(pairs, sourceLang, targetLang);

    try {
        await Word.run(async (context) => {
            if (memorySettings.replaceAll) {
                // Replace ALL occurrences in the whole document
                let replacedCount = 0;
                for (const [src, tgt] of Object.entries(pairs)) {
                    if (!src.trim()) continue;
                    const searchResults = context.document.body.search(src.trim(), { matchCase: false, matchWholeWord: false });
                    searchResults.load('items');
                    await context.sync();
                    for (const item of searchResults.items) {
                        item.insertText(tgt, Word.InsertLocation.replace);
                        replacedCount++;
                    }
                    await context.sync();
                }
                showStatus(`✅ Replaced ${replacedCount} occurrence(s) in document`, 'success');
            } else {
                // Replace selected text only
                const range = context.document.getSelection();
                range.load('text');
                await context.sync();

                const selectedText = range.text.trim();
                if (!selectedText) {
                    showStatus('⚠️ No text selected. Select text in Word first.', 'error');
                    return;
                }

                // Find best match in pairs
                const match = findBestPairMatch(selectedText, pairs);
                if (match) {
                    range.insertText(match, Word.InsertLocation.replace);
                    await context.sync();
                    recordCacheHit(model);
                    showStatus(`✅ Applied: "${selectedText}" → "${match}"`, 'success');
                } else {
                    showStatus(`⚠️ No match found for "${selectedText}" in chat response`, 'error');
                }
            }
        });
    } catch (err) {
        showStatus(`❌ Error: ${err.message}`, 'error');
    }
}

// Parse "source | translation" lines from chat response
function parseChatPairs(text) {
    const pairs = {};
    const lines = text.split('\n');
    for (const line of lines) {
        if (line.includes('|')) {
            const parts = line.split('|');
            if (parts.length >= 2) {
                // Clean markdown bold/asterisks and extra spaces
                const src = parts[0].replace(/\*\*/g, '').replace(/^[-•\s]+/, '').trim();
                const tgt = parts[1].replace(/\*\*/g, '').trim();
                if (src && tgt) pairs[src] = tgt;
            }
        }
    }
    return pairs;
}

// Find exact or close match for selected text in pairs
function findBestPairMatch(selectedText, pairs) {
    const normalized = selectedText.trim().toLowerCase();
    // 1. Exact match
    for (const [src, tgt] of Object.entries(pairs)) {
        if (src.trim().toLowerCase() === normalized) return tgt;
    }
    // 2. Contains match (selected text is inside a source key)
    for (const [src, tgt] of Object.entries(pairs)) {
        if (src.trim().toLowerCase().includes(normalized) || normalized.includes(src.trim().toLowerCase())) return tgt;
    }
    return null;
}

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("✅ ChatGPT Translator Ready!");
        
        // Load saved API Key
        const savedKey = localStorage.getItem('openai_api_key');
        if (savedKey) {
            document.getElementById('apiKey').value = savedKey;
            document.getElementById('apiKeySaved').style.display = 'inline-block';
        }
        
        // Load saved Google Sheets URL
        const savedSheetUrl = localStorage.getItem('google_sheet_url');
        if (savedSheetUrl) {
            document.getElementById('googleSheetUrl').value = savedSheetUrl;
            document.getElementById('sheetSaved').style.display = 'inline-block';
        }
        
        // Load usage stats
        loadUsageStats();
        
        // Load translation memory system
        loadMemorySystem();
        
        // Update model display
        updateCurrentModel();
        
        // Listen for model changes
        document.getElementById('modelSelect').addEventListener('change', updateCurrentModel);
    }
});

// Save API Key
function saveApiKey() {
    const apiKey = document.getElementById('apiKey').value.trim();
    
    if (!apiKey) {
        showStatus('❌ Please enter an API Key', 'error');
        return;
    }
    
    if (!apiKey.startsWith('sk-')) {
        showStatus('❌ Invalid API Key. Must start with sk-', 'error');
        return;
    }
    
    localStorage.setItem('openai_api_key', apiKey);
    document.getElementById('apiKeySaved').style.display = 'inline-block';
    showStatus('✅ API Key saved successfully', 'success');
}

// Save Google Sheet URL
function saveGoogleSheetUrl() {
    const sheetUrl = document.getElementById('googleSheetUrl').value.trim();
    
    if (!sheetUrl) {
        localStorage.removeItem('google_sheet_url');
        document.getElementById('sheetSaved').style.display = 'none';
        showStatus('📊 Google Sheets logging disabled', 'info');
        return;
    }
    
    if (!sheetUrl.includes('script.google.com')) {
        showStatus('⚠️ Please enter a valid Google Apps Script URL', 'error');
        return;
    }
    
    localStorage.setItem('google_sheet_url', sheetUrl);
    document.getElementById('sheetSaved').style.display = 'inline-block';
    showStatus('✅ Google Sheets connected! All translations will be logged.', 'success');
}

// Show Google Sheets setup instructions
function showSheetInstructions() {
    const instructions = `📊 Google Sheets Setup Guide:

1. Create a new Google Sheet at sheets.google.com

2. Add headers in Row 1:
   Timestamp | Source Lang | Target Lang | Mode | Model | Words | Input Tokens | Output Tokens | Cost (USD) | Text Preview | User

3. Go to Extensions → Apps Script

4. Paste this code:
   function doPost(e) {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     const data = JSON.parse(e.postData.contents);
     sheet.appendRow([new Date(), data.sourceLang, data.targetLang, data.mode, data.model, data.words, data.inputTokens, data.outputTokens, data.cost, data.textPreview, data.user]);
     return ContentService.createTextOutput('OK');
   }

5. Deploy → New deployment → Web app
   - Execute as: Me
   - Who has access: Anyone
   - Copy the deployment URL

6. Paste the URL here and click Save Sheet URL!

Done! All translations will log automatically 🎉`;
    
    alert(instructions);
}

// Main translation function
async function translateSelection() {
    const apiKey = localStorage.getItem('openai_api_key');
    
    if (!apiKey) {
        showStatus('❌ Please enter and save your API Key first', 'error');
        return;
    }
    
    showStatus('⏳ Translating...', 'info');
    
    try {
        await Word.run(async (context) => {
            // Get selected range
            const range = context.document.getSelection();
            range.load('text, paragraphs');
            await context.sync();
            
            const selectedText = range.text;
            
            if (!selectedText || selectedText.trim().length === 0) {
                showStatus('❌ No text selected. Please select text to translate.', 'error');
                return;
            }
            
            // Get language settings
            const sourceLang = document.getElementById('sourceLang').value;
            const targetLang = document.getElementById('targetLang').value;
            const model = document.getElementById('modelSelect').value;
            
            if (sourceLang === targetLang && sourceLang !== 'auto') {
                showStatus('⚠️ Source and target languages are the same', 'error');
                return;
            }
            
            // Get all paragraphs in selection
            const paragraphs = range.paragraphs;
            paragraphs.load('items');
            await context.sync();
            
            // If only one paragraph, translate normally
            if (paragraphs.items.length === 1) {
                await translateSingleParagraph(context, paragraphs.items[0], sourceLang, targetLang, apiKey, model);
            } else {
                // Multiple paragraphs - translate each separately to preserve formatting
                await translateMultipleParagraphs(context, paragraphs.items, sourceLang, targetLang, apiKey, model);
            }
        });
    } catch (error) {
        console.error('Translation error:', error);
        showStatus(`❌ Error: ${error.message}`, 'error');
    }
}

// Translate a single paragraph
async function translateSingleParagraph(context, paragraph, sourceLang, targetLang, apiKey, model) {
    paragraph.load('text, font, style');
    await context.sync();
    
    const text = paragraph.text.trim();
    
    if (!text || text.length === 0) {
        showStatus('✅ Translation completed!', 'success');
        return;
    }
    
    // Save original formatting (with safety checks)
    const originalFont = {
        name: paragraph.font.name || 'Calibri',
        size: paragraph.font.size || 11,
        bold: paragraph.font.bold || false,
        italic: paragraph.font.italic || false,
        underline: paragraph.font.underline || 'None',
        color: paragraph.font.color || '#000000'
    };
    
    const wordCount = text.split(/\s+/).length;
    
    // Check translation memory cache first (exact match only)
    const cached = lookupCache(text, sourceLang, targetLang);
    if (cached) {
        const model = document.getElementById('modelSelect').value;
        paragraph.clear();
        paragraph.insertText(cached, Word.InsertLocation.start);
        await context.sync();
        paragraph.font.name = originalFont.name;
        paragraph.font.size = originalFont.size;
        paragraph.font.bold = originalFont.bold;
        paragraph.font.italic = originalFont.italic;
        paragraph.font.underline = originalFont.underline;
        paragraph.font.color = originalFont.color;
        await context.sync();
        recordCacheHit(model);
        showStatus(`⚡ From memory: "${text.substring(0, 30)}..." → "${cached.substring(0, 30)}..."`, 'success');
        return;
    }
    
    const result = await callChatGPT(text, sourceLang, targetLang, apiKey, model);
    
    // Clear and insert new text
    paragraph.clear();
    paragraph.insertText(result.translation, Word.InsertLocation.start);
    await context.sync();
    
    // Restore formatting
    paragraph.font.name = originalFont.name;
    paragraph.font.size = originalFont.size;
    paragraph.font.bold = originalFont.bold;
    paragraph.font.italic = originalFont.italic;
    paragraph.font.underline = originalFont.underline;
    paragraph.font.color = originalFont.color;
    
    await context.sync();
    
    updateUsageStats(wordCount, result.usage, model, result.detectedLanguage);
    
    // Save to cache for future use
    saveToCache(text, result.translation, sourceLang, targetLang);
    
    const detectedMsg = result.detectedLanguage ? ` (Detected: ${result.detectedLanguage})` : '';
    showStatus(`✅ Translation completed!${detectedMsg}`, 'success');
}

// Translate multiple paragraphs while preserving individual formatting
async function translateMultipleParagraphs(context, paragraphs, sourceLang, targetLang, apiKey, model) {
    let totalWords = 0;
    let totalUsage = { prompt_tokens: 0, completion_tokens: 0 };
    let translatedCount = 0;
    
    // Collect all text first
    const paragraphsData = [];
    
    for (const para of paragraphs) {
        para.load('text, font, style');
        await context.sync();
        
        const text = para.text.trim();
        
        if (text && text.length > 0) {
            paragraphsData.push({
                paragraph: para,
                text: text,
                font: {
                    name: para.font.name || 'Calibri',
                    size: para.font.size || 11,
                    bold: para.font.bold || false,
                    italic: para.font.italic || false,
                    underline: para.font.underline || 'None',
                    color: para.font.color || '#000000'
                }
            });
        }
    }
    
    if (paragraphsData.length === 0) {
        showStatus('❌ No text to translate', 'error');
        return;
    }
    
    showStatus(`⏳ Translating ${paragraphsData.length} paragraphs...`, 'info');
    
    // Use numbered markers to reliably split segments back after translation
    // This avoids issues with ChatGPT not preserving exact \n\n separators (especially in tables)
    const markedText = paragraphsData.map((p, i) => `[SEG:${i + 1}]\n${p.text}`).join('\n\n');
    const wordCount = markedText.split(/\s+/).length;
    
    try {
        const result = await callChatGPT(markedText, sourceLang, targetLang, apiKey, model, true);
        
        // Parse segments by markers [SEG:N]
        const translations = parseSegments(result.translation, paragraphsData.length);
        
        // Apply translations back to paragraphs
        for (let i = 0; i < paragraphsData.length && i < translations.length; i++) {
            const data = paragraphsData[i];
            
            // Clear and insert translated text
            data.paragraph.clear();
            data.paragraph.insertText(translations[i], Word.InsertLocation.start);
            await context.sync();
            
            // Restore original formatting for this specific paragraph
            data.paragraph.font.name = data.font.name;
            data.paragraph.font.size = data.font.size;
            data.paragraph.font.bold = data.font.bold;
            data.paragraph.font.italic = data.font.italic;
            data.paragraph.font.underline = data.font.underline;
            data.paragraph.font.color = data.font.color;
            
            await context.sync();
            
            translatedCount++;
        }
        
        updateUsageStats(wordCount, result.usage, model, result.detectedLanguage);
        
        showStatus(`✅ Translated ${translatedCount} paragraphs successfully!`, 'success');
        
    } catch (error) {
        console.error('Translation error:', error);
        showStatus(`❌ Error: ${error.message}`, 'error');
    }
}

// Parse numbered segment markers [SEG:N] from translation response
function parseSegments(text, expectedCount) {
    const results = [];
    // Match each [SEG:N] marker and capture text until next marker or end
    const regex = /\[SEG:(\d+)\]\s*([\s\S]*?)(?=\[SEG:\d+\]|$)/g;
    let match;
    const map = {};

    while ((match = regex.exec(text)) !== null) {
        const idx = parseInt(match[1]) - 1; // 0-based
        map[idx] = match[2].trim();
    }

    for (let i = 0; i < expectedCount; i++) {
        results.push(map[i] || ''); // fallback to empty if missing
    }

    // Fallback: if no markers found, split by \n\n like before
    if (Object.keys(map).length === 0) {
        const fallback = text.split('\n\n').filter(t => t.trim());
        for (let i = 0; i < expectedCount; i++) {
            results[i] = fallback[i] || results[i] || '';
        }
    }

    return results;
}

// Call ChatGPT API
async function callChatGPT(text, fromLang, toLang, apiKey, model, useSegmentMarkers = false) {
    const targetLanguage = LANGUAGE_MAP[toLang];
    const mode = document.getElementById('translationMode').value;
    const style = document.getElementById('translationStyle').value;
    
    // Build system prompt based on mode and style
    let systemPrompt = buildSystemPrompt(fromLang, toLang, targetLanguage, mode, style);
    
    // Lock numbers instruction
    if (memorySettings.lockNumbers) {
        systemPrompt += `\n\nCRITICAL - LOCK RULE: Never modify, translate, or alter any numbers, digits, dates, file IDs, reference numbers, phone numbers, or legal article numbers. Keep them exactly as they appear in the source text. Example: "المادة 15" → "Article 15" (not "Article fifteen").`;
    }
    
    // When using segment markers, instruct the model to preserve them exactly
    if (useSegmentMarkers) {
        systemPrompt += `\n\nCRITICAL: The text contains segment markers like [SEG:1], [SEG:2], etc. You MUST preserve these markers EXACTLY in your output, each on its own line before its corresponding translated segment. Do NOT remove, renumber, or merge any markers.`;
    }
    
    // Adjust temperature based on mode
    const temperature = getTemperatureForMode(mode, style);
    
    // Use max_completion_tokens for newer models (gpt-5.2, o1), max_tokens for older ones
    const isNewerModel = model.includes('gpt-5') || model.includes('o1');
    const tokenParam = isNewerModel ? 'max_completion_tokens' : 'max_tokens';
    
    const requestBody = {
        model: model,
        messages: [
            {
                role: 'system',
                content: systemPrompt
            },
            {
                role: 'user',
                content: text
            }
        ],
        temperature: temperature
    };
    
    // Add the appropriate token parameter
    requestBody[tokenParam] = 3000;
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error?.message || `HTTP Error: ${response.status}`);
        }
        
        const data = await response.json();
        
        if (!data.choices || !data.choices[0]) {
            throw new Error('Unexpected API response');
        }
        
        let translation = data.choices[0].message.content.trim();
        let detectedLanguage = null;
        
        // Extract detected language if auto-detect was used
        if (fromLang === 'auto' && translation.includes('DETECTED:')) {
            const parts = translation.split('DETECTED:');
            translation = parts[0].trim();
            detectedLanguage = parts[1].trim();
        }
        
        return {
            translation: translation,
            usage: data.usage,
            detectedLanguage: detectedLanguage
        };
        
    } catch (error) {
        if (error.message.includes('fetch')) {
            throw new Error('Failed to connect to ChatGPT. Check your internet.');
        } else if (error.message.includes('Incorrect API key')) {
            throw new Error('Invalid API Key. Please check your key.');
        } else if (error.message.includes('quota') || error.message.includes('insufficient_quota')) {
            throw new Error('API quota exceeded. Add credits to your OpenAI account.');
        }
        throw error;
    }
}

// Update usage statistics
function updateUsageStats(wordCount, usage, model, detectedLanguage) {
    // Safety check: make sure usage object exists
    if (!usage) {
        console.warn('No usage data returned from API');
        return;
    }
    
    usageStats.totalTranslations++;
    usageStats.totalWords += wordCount;
    usageStats.totalInputTokens += usage.prompt_tokens || 0;
    usageStats.totalOutputTokens += usage.completion_tokens || 0;
    
    // Calculate cost with safety check for pricing
    const pricing = MODEL_PRICING[model];
    let currentCost = 0;
    if (pricing) {
        const inputCost = ((usage.prompt_tokens || 0) / 1000000) * pricing.input;
        const outputCost = ((usage.completion_tokens || 0) / 1000000) * pricing.output;
        currentCost = inputCost + outputCost;
        usageStats.totalCost += currentCost;
    } else {
        console.warn(`No pricing data for model: ${model}`);
    }
    
    // Save to localStorage
    localStorage.setItem('usage_stats', JSON.stringify(usageStats));
    
    const logEntry = {
        sourceLang: document.getElementById('sourceLang')?.value || 'auto',
        targetLang: document.getElementById('targetLang')?.value || 'ar',
        mode: document.getElementById('translationMode')?.value || 'general',
        model: model,
        words: wordCount,
        inputTokens: usage.prompt_tokens || 0,
        outputTokens: usage.completion_tokens || 0,
        cost: currentCost.toFixed(6)
    };

    // Log to server (persistent dashboard)
    logToServer(logEntry);
    syncToServer();

    // Log to Google Sheets (if configured)
    logToGoogleSheets({ ...logEntry, textPreview: '', user: localStorage.getItem('user_name') || 'Anonymous' });
    
    // Update display
    displayUsageStats();
}

// Log translation data to Google Sheets
async function logToGoogleSheets(data) {
    const sheetUrl = localStorage.getItem('google_sheet_url');
    
    if (!sheetUrl) {
        // Silently skip if not configured
        return;
    }
    
    try {
        await fetch(sheetUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });
        // Note: no-cors mode means we can't read the response, but that's ok
        console.log('Logged to Google Sheets');
    } catch (error) {
        console.error('Failed to log to Google Sheets:', error);
    }
}

// Load usage statistics
function loadUsageStats() {
    const saved = localStorage.getItem('usage_stats');
    if (saved) {
        usageStats = JSON.parse(saved);
    }
    displayUsageStats();
}

// Display usage statistics
function displayUsageStats() {
    document.getElementById('totalTranslations').textContent = usageStats.totalTranslations.toLocaleString();
    document.getElementById('totalWords').textContent = usageStats.totalWords.toLocaleString();
    document.getElementById('totalInputTokens').textContent = usageStats.totalInputTokens.toLocaleString();
    document.getElementById('totalOutputTokens').textContent = usageStats.totalOutputTokens.toLocaleString();
    document.getElementById('totalCost').textContent = `$${usageStats.totalCost.toFixed(4)}`;
}

// Update current model display
function updateCurrentModel() {
    const model = document.getElementById('modelSelect').value;
    document.getElementById('currentModel').textContent = model;
}

// Reset statistics
function resetStats() {
    if (confirm('Are you sure you want to reset all usage statistics?')) {
        usageStats = {
            totalTranslations: 0,
            totalWords: 0,
            totalInputTokens: 0,
            totalOutputTokens: 0,
            totalCost: 0
        };
        localStorage.setItem('usage_stats', JSON.stringify(usageStats));
        displayUsageStats();
        showStatus('✅ Statistics reset successfully', 'success');
    }
}

// Show status messages
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.style.display = 'block';
    
    // Auto-hide success messages
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}

// ===== CHAT FUNCTIONALITY =====

// Chat history (memory)
let chatHistory = [];

// Add message to chat UI
function addChatBubble(role, text) {
    const container = document.getElementById('chatMessages');

    // Remove welcome message on first message
    const welcome = container.querySelector('.chat-welcome');
    if (welcome) welcome.remove();

    const bubble = document.createElement('div');

    if (role === 'user') {
        bubble.className = 'chat-bubble user';
        bubble.textContent = text;
    } else if (role === 'system-msg') {
        bubble.className = 'chat-bubble system-msg';
        bubble.textContent = text;
    } else {
        bubble.className = 'chat-bubble assistant';
        bubble.innerHTML = `<div class="bubble-label">🤖 Assistant</div>${text.replace(/\n/g, '<br>')}`;
    }

    container.appendChild(bubble);
    container.scrollTop = container.scrollHeight;
    return bubble;
}

// Show typing indicator
function showTyping() {
    const container = document.getElementById('chatMessages');
    const typing = document.createElement('div');
    typing.className = 'typing-indicator';
    typing.id = 'typingIndicator';
    typing.innerHTML = '<div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div>';
    container.appendChild(typing);
    container.scrollTop = container.scrollHeight;
}

// Remove typing indicator
function hideTyping() {
    const typing = document.getElementById('typingIndicator');
    if (typing) typing.remove();
}

// Handle Enter key in chat input
function handleChatKeydown(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        sendChatMessage();
    }
}

// Send message from quick action buttons
async function sendQuickMessage(action) {
    const apiKey = localStorage.getItem('openai_api_key');
    if (!apiKey) {
        addChatBubble('system-msg', '⚠️ Please save your API key first!');
        return;
    }

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load('text');
            await context.sync();

            const selectedText = range.text.trim();

            if (!selectedText) {
                addChatBubble('system-msg', '⚠️ Please select text in your document first!');
                return;
            }

            let userMessage = '';
            if (action === 'translate_selection') {
                const targetLang = LANGUAGE_MAP[document.getElementById('targetLang').value];
                userMessage = `Please translate this text to ${targetLang}:\n\n"${selectedText}"`;
            } else if (action === 'explain_selection') {
                userMessage = `Please explain what this text means:\n\n"${selectedText}"`;
            } else if (action === 'improve_selection') {
                userMessage = `Please improve and refine this text:\n\n"${selectedText}"`;
            }

            await processChatMessage(userMessage, apiKey);
        });
    } catch (e) {
        addChatBubble('system-msg', '⚠️ Could not read selection from document.');
        const apiKey = localStorage.getItem('openai_api_key');
        if (action === 'translate_selection') {
            await processChatMessage('Please help me translate some text.', apiKey);
        }
    }
}

// Main chat send function
async function sendChatMessage() {
    const input = document.getElementById('chatInput');
    const userMessage = input.value.trim();

    if (!userMessage) return;

    const apiKey = localStorage.getItem('openai_api_key');
    if (!apiKey) {
        addChatBubble('system-msg', '⚠️ Please save your API key first!');
        return;
    }

    input.value = '';
    input.style.height = 'auto';

    await processChatMessage(userMessage, apiKey);
}

// Process and send chat message to API
async function processChatMessage(userMessage, apiKey) {
    const sendBtn = document.getElementById('sendBtn');
    sendBtn.disabled = true;

    // Add user message to UI and history
    addChatBubble('user', userMessage);
    chatHistory.push({ role: 'user', content: userMessage });

    // Show typing
    showTyping();

    try {
        const model = document.getElementById('modelSelect').value;

        // Build messages array with system context
        const messages = [
            {
                role: 'system',
                content: `You are an expert translation assistant for Dar Al Marjaan Translation Services. 
You help with document translation, terminology, and language questions.
You are working inside Microsoft Word as an add-in.
When the user gives you context about a document (names, IDs, terminology), remember it for future translations.
Be concise but helpful. If asked to translate text, provide the translation directly.
IMPORTANT: When translating a list of terms or phrases, format each pair as:
source text | translated text
(one per line, pipe-separated) so the user can apply them directly to Word.
${memorySettings.lockNumbers ? 'CRITICAL: Never modify numbers, dates, IDs, or reference numbers.' : ''}
Current translation settings: Mode = ${document.getElementById('translationMode')?.value || 'general'}, Style = ${document.getElementById('translationStyle')?.value || 'balanced'}.`
            },
            ...chatHistory
        ];

        // Prepare request
        const requestBody = {
            model: model,
            messages: messages,
            temperature: 0.5
        };

        if (model.includes('gpt-5') || model.includes('o1')) {
            requestBody.max_completion_tokens = 1000;
        } else {
            requestBody.max_tokens = 1000;
        }

        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });

        hideTyping();

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error?.message || `HTTP Error: ${response.status}`);
        }

        const data = await response.json();
        const assistantMessage = data.content?.[0]?.text || data.choices?.[0]?.message?.content || 'No response received.';

        // Add to history and UI
        chatHistory.push({ role: 'assistant', content: assistantMessage });
        addChatBubble('assistant', assistantMessage);

        // Keep history manageable (last 20 messages)
        if (chatHistory.length > 20) {
            chatHistory = chatHistory.slice(-20);
        }

    } catch (error) {
        hideTyping();
        addChatBubble('system-msg', `❌ Error: ${error.message}`);
        // Remove last user message from history on error
        chatHistory.pop();
    } finally {
        sendBtn.disabled = false;
    }
}

// Clear chat history
function clearChat() {
    chatHistory = [];
    const container = document.getElementById('chatMessages');
    container.innerHTML = `
        <div class="chat-welcome">
            <div class="chat-welcome-icon">🤖</div>
            <div><strong>Dar Al Marjaan Assistant</strong></div>
            <div style="margin-top:6px;">Ask me anything about your document, give me context about your project, or request a translation!</div>
        </div>
    `;
}

// ===== FILE UPLOAD & IMAGE TRANSLATION =====

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const apiKey = localStorage.getItem('openai_api_key');
    if (!apiKey) {
        addChatBubble('system-msg', '⚠️ Please save your API key first!');
        return;
    }

    const uploadStatus = document.getElementById('uploadStatus');
    const uploadBtn = document.querySelector('.btn-upload');
    const targetLang = LANGUAGE_MAP[document.getElementById('targetLang').value] || 'Arabic';
    const mode = document.getElementById('translationMode').value;

    // Reset file input
    event.target.value = '';

    // Show loading
    uploadBtn.disabled = true;
    uploadBtn.textContent = '⏳ Processing...';
    uploadStatus.style.display = 'block';
    uploadStatus.textContent = `📎 Reading: ${file.name}`;

    try {
        if (file.type === 'application/pdf') {
            // PDF handling - convert first page to image
            await handlePDFUpload(file, apiKey, targetLang, mode, uploadStatus);
        } else if (file.type.startsWith('image/')) {
            // Image handling
            await handleImageUpload(file, apiKey, targetLang, mode, uploadStatus);
        } else {
            addChatBubble('system-msg', '❌ Please upload an image (JPG, PNG) or PDF file.');
        }
    } catch (error) {
        addChatBubble('system-msg', `❌ Error: ${error.message}`);
    } finally {
        uploadBtn.disabled = false;
        uploadBtn.textContent = '📎 Upload Image / PDF → Translate';
        uploadStatus.style.display = 'none';
    }
}

// Handle image file upload
async function handleImageUpload(file, apiKey, targetLang, mode, uploadStatus) {
    uploadStatus.textContent = '🔍 Reading image with AI...';

    // Convert to base64
    const base64 = await fileToBase64(file);
    const mediaType = file.type;

    addChatBubble('user', `📎 Uploaded: ${file.name}\n🌐 Translate to: ${targetLang}`);

    showTyping();
    uploadStatus.textContent = '🤖 AI is translating...';

    const model = document.getElementById('modelSelect').value;
    const modePrompt = getModePromptForUpload(mode, targetLang);

    const requestBody = {
        model: 'gpt-4o',  // Always use gpt-4o for vision (supports images)
        messages: [
            ...chatHistory,
            {
                role: 'user',
                content: [
                    {
                        type: 'image_url',
                        image_url: {
                            url: `data:${mediaType};base64,${base64}`,
                            detail: 'high'
                        }
                    },
                    {
                        type: 'text',
                        text: modePrompt
                    }
                ]
            }
        ],
        max_tokens: 4000
    };

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify(requestBody)
    });

    hideTyping();

    if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error?.message || `HTTP Error: ${response.status}`);
    }

    const data = await response.json();
    const result = data.choices?.[0]?.message?.content || 'No response received.';

    // Add to chat history
    chatHistory.push({
        role: 'user',
        content: `[Image uploaded: ${file.name}] Please translate to ${targetLang}`
    });
    chatHistory.push({ role: 'assistant', content: result });

    addChatBubble('assistant', result);

    // Update usage stats
    if (data.usage) {
        updateUsageStats(
            result.split(/\s+/).length,
            data.usage,
            'gpt-4o',
            null
        );
    }
}

// Handle PDF upload - convert to image using canvas then send
async function handlePDFUpload(file, apiKey, targetLang, mode, uploadStatus) {
    uploadStatus.textContent = '📄 Converting PDF pages to images...';

    try {
        addChatBubble('user', `📎 Uploaded PDF: ${file.name}\n🌐 Translate to: ${targetLang}`);

        // Load PDF.js from CDN
        if (!window.pdfjsLib) {
            uploadStatus.textContent = '⏳ Loading PDF engine...';
            await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }

        // Read PDF file
        const arrayBuffer = await file.arrayBuffer();
        const pdfDoc = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        const totalPages = pdfDoc.numPages;
        const pageImages = [];

        uploadStatus.textContent = `📄 Converting ${totalPages} page(s)...`;

        // Convert each page to image (max 6 pages to avoid token limits)
        const maxPages = Math.min(totalPages, 6);

        for (let pageNum = 1; pageNum <= maxPages; pageNum++) {
            uploadStatus.textContent = `📄 Converting page ${pageNum}/${maxPages}...`;

            const page = await pdfDoc.getPage(pageNum);
            const viewport = page.getViewport({ scale: 2.0 }); // High quality

            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext('2d');

            await page.render({ canvasContext: ctx, viewport }).promise;

            const base64 = canvas.toDataURL('image/jpeg', 0.85).split(',')[1];
            pageImages.push(base64);
        }

        if (totalPages > maxPages) {
            addChatBubble('system-msg', `📄 PDF has ${totalPages} pages. Translating first ${maxPages} pages.`);
        }

        showTyping();
        uploadStatus.textContent = '🤖 AI is reading and translating...';

        const modePrompt = getModePromptForUpload(mode, targetLang);

        // Build content array with all page images
        const contentArray = [{ type: 'text', text: modePrompt }];

        for (let i = 0; i < pageImages.length; i++) {
            contentArray.push({
                type: 'text',
                text: pageImages.length > 1 ? `--- Page ${i + 1} ---` : ''
            });
            contentArray.push({
                type: 'image_url',
                image_url: {
                    url: `data:image/jpeg;base64,${pageImages[i]}`,
                    detail: 'high'
                }
            });
        }

        const requestBody = {
            model: 'gpt-4o',
            messages: [
                ...chatHistory,
                { role: 'user', content: contentArray }
            ],
            max_tokens: 4000
        };

        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });

        hideTyping();

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error?.message || `HTTP Error: ${response.status}`);
        }

        const data = await response.json();
        const result = data.choices?.[0]?.message?.content || 'No response received.';

        chatHistory.push({ role: 'user', content: `[PDF uploaded: ${file.name}, ${maxPages} pages]` });
        chatHistory.push({ role: 'assistant', content: result });

        addChatBubble('assistant', result);

        if (data.usage) {
            updateUsageStats(result.split(/\s+/).length, data.usage, 'gpt-4o', null);
        }

    } catch (error) {
        hideTyping();
        // If PDF.js fails, ask user to take screenshot
        if (error.message.includes('PDF') || error.message.includes('script')) {
            addChatBubble('system-msg', '💡 Could not process PDF automatically. Please take a screenshot of the page (Windows + Shift + S) and upload as image instead.');
        } else {
            throw error;
        }
    }
}

// Load external script dynamically
function loadScript(src) {
    return new Promise((resolve, reject) => {
        if (document.querySelector(`script[src="${src}"]`)) {
            resolve();
            return;
        }
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

// Build translation prompt based on mode
function getModePromptForUpload(mode, targetLang) {
    const modeInstructions = {
        'legal': `You are a certified legal translator. Extract ALL text from this document and translate it to ${targetLang}. Use precise legal terminology. Keep all names, dates, reference numbers, and IDs exactly as they appear.`,
        'certificate': `You are an official document translator. Extract ALL text from this certificate and translate it to ${targetLang}. Keep names, dates, ID numbers, and official seals unchanged. Use formal government language.`,
        'bank': `You are a financial document translator. Extract ALL text from this bank document and translate it to ${targetLang}. Keep account numbers, amounts, dates, and reference codes unchanged.`,
        'medical': `You are a medical translator. Extract ALL text from this medical document and translate it to ${targetLang}. Keep patient names, dates, test values, and measurements unchanged.`,
        'government': `You are an official government document translator. Extract ALL text and translate to ${targetLang}. Keep all reference numbers, dates, names, and official codes unchanged.`,
        'business': `You are a business document translator. Extract ALL text and translate to ${targetLang}. Keep company names, amounts, dates, and clause numbers unchanged.`
    };

    const basePrompt = modeInstructions[mode] || `You are a professional translator. Extract ALL text from this document and translate it to ${targetLang}. Keep all names, dates, numbers, and IDs exactly as they appear.`;

    return `${basePrompt}

Format your response as:
**ORIGINAL TEXT (extracted):**
[the extracted text]

**TRANSLATION (${targetLang}):**
[the full translation]

CRITICAL RULES:
- Extract EVERY word visible in the document
- Do NOT skip any text
- Preserve the document structure and layout in the translation
- Keep all numbers, dates, names, and IDs unchanged`;
}

// Convert file to base64
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}
