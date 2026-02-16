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
    'gpt-5.2': { input: 20.00, output: 80.00 },     // Premium for complex files (estimated)
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
        return `${basePrompt} Detect the source language and translate to ${targetLanguage}. IMPORTANT: Preserve all numbers, dates, IDs, and proper names exactly as they appear. Return ONLY the translated text without any explanations. At the very end, on a new line, write "DETECTED:" followed by the detected language name in English.`;
    } else {
        const sourceLanguage = LANGUAGE_MAP[fromLang];
        return `${basePrompt} Translate from ${sourceLanguage} to ${targetLanguage}. IMPORTANT: Preserve all numbers, dates, IDs, and proper names exactly as they appear. Return ONLY the translated text without any explanations.`;
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
        
        // Load usage stats
        loadUsageStats();
        
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

// Main translation function
async function translateSelection() {
    const apiKey = localStorage.getItem('openai_api_key');
    
    if (!apiKey) {
        showStatus('❌ Please enter and save your API Key first', 'error');
        return;
    }
    
    showStatus('⏳ Analyzing selection...', 'info');
    
    try {
        await Word.run(async (context) => {
            // Get selected range
            const range = context.document.getSelection();
            range.load('text, type, paragraphs, tables');
            await context.sync();
            
            // Check if selection contains tables
            const tables = range.tables;
            tables.load('items');
            await context.sync();
            
            // Get language settings
            const sourceLang = document.getElementById('sourceLang').value;
            const targetLang = document.getElementById('targetLang').value;
            const model = document.getElementById('modelSelect').value;
            
            if (sourceLang === targetLang && sourceLang !== 'auto') {
                showStatus('⚠️ Source and target languages are the same', 'error');
                return;
            }
            
            // If we have tables in selection
            if (tables.items.length > 0) {
                showStatus('⏳ Translating table cells...', 'info');
                await translateTableCells(context, tables.items, sourceLang, targetLang, apiKey, model);
            } else {
                // Regular text translation (original behavior)
                const selectedText = range.text;
                
                if (!selectedText || selectedText.trim().length === 0) {
                    showStatus('❌ No text selected. Please select text to translate.', 'error');
                    return;
                }
                
                showStatus('⏳ Translating...', 'info');
                
                const wordCount = selectedText.trim().split(/\s+/).length;
                const result = await callChatGPT(selectedText, sourceLang, targetLang, apiKey, model);
                
                range.insertText(result.translation, Word.InsertLocation.replace);
                await context.sync();
                
                updateUsageStats(wordCount, result.usage, model, result.detectedLanguage);
                
                const detectedMsg = result.detectedLanguage ? ` (Detected: ${result.detectedLanguage})` : '';
                showStatus(`✅ Translation completed!${detectedMsg}`, 'success');
            }
        });
    } catch (error) {
        console.error('Translation error:', error);
        showStatus(`❌ Error: ${error.message}`, 'error');
    }
}

// Translate table cells
async function translateTableCells(context, tables, sourceLang, targetLang, apiKey, model) {
    let totalCells = 0;
    let translatedCells = 0;
    let totalWords = 0;
    let totalUsage = { prompt_tokens: 0, completion_tokens: 0 };
    
    try {
        // Process each table
        for (const table of tables) {
            table.load('rowCount, rows');
            await context.sync();
            
            const rows = table.rows;
            rows.load('items');
            await context.sync();
            
            // Process each row
            for (const row of rows.items) {
                row.load('cells');
                await context.sync();
                
                const cells = row.cells;
                cells.load('items');
                await context.sync();
                
                // Process each cell
                for (const cell of cells.items) {
                    cell.load('body');
                    await context.sync();
                    
                    const cellBody = cell.body;
                    cellBody.load('text');
                    await context.sync();
                    
                    const cellText = cellBody.text.trim();
                    
                    // Skip empty cells
                    if (!cellText || cellText.length === 0) {
                        continue;
                    }
                    
                    totalCells++;
                    
                    // Update status every 10 cells
                    if (totalCells % 10 === 0) {
                        showStatus(`⏳ Translating cell ${totalCells}...`, 'info');
                    }
                    
                    try {
                        // Translate the cell
                        const wordCount = cellText.split(/\s+/).length;
                        const result = await callChatGPT(cellText, sourceLang, targetLang, apiKey, model);
                        
                        // Replace cell content
                        cellBody.clear();
                        cellBody.insertText(result.translation, Word.InsertLocation.start);
                        await context.sync();
                        
                        translatedCells++;
                        totalWords += wordCount;
                        
                        // Accumulate usage stats
                        if (result.usage) {
                            totalUsage.prompt_tokens += result.usage.prompt_tokens || 0;
                            totalUsage.completion_tokens += result.usage.completion_tokens || 0;
                        }
                        
                    } catch (cellError) {
                        console.error(`Error translating cell:`, cellError);
                        // Continue with next cell even if one fails
                    }
                }
            }
        }
        
        // Update total usage stats
        updateUsageStats(totalWords, totalUsage, model, null);
        
        showStatus(`✅ Translated ${translatedCells} cells successfully!`, 'success');
        
    } catch (error) {
        console.error('Table translation error:', error);
        showStatus(`❌ Error: Translated ${translatedCells}/${totalCells} cells. ${error.message}`, 'error');
    }
}

// Call ChatGPT API
async function callChatGPT(text, fromLang, toLang, apiKey, model) {
    const targetLanguage = LANGUAGE_MAP[toLang];
    const mode = document.getElementById('translationMode').value;
    const style = document.getElementById('translationStyle').value;
    
    // Build system prompt based on mode and style
    let systemPrompt = buildSystemPrompt(fromLang, toLang, targetLanguage, mode, style);
    
    // Adjust temperature based on mode
    const temperature = getTemperatureForMode(mode, style);
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
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
                temperature: temperature,
                max_tokens: 3000
            })
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
    if (pricing) {
        const inputCost = ((usage.prompt_tokens || 0) / 1000000) * pricing.input;
        const outputCost = ((usage.completion_tokens || 0) / 1000000) * pricing.output;
        usageStats.totalCost += (inputCost + outputCost);
    } else {
        console.warn(`No pricing data for model: ${model}`);
    }
    
    // Save to localStorage
    localStorage.setItem('usage_stats', JSON.stringify(usageStats));
    
    // Update display
    displayUsageStats();
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
