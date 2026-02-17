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
    'gpt-4o-mini': { input: 0.150, output: 0.600 },
    'gpt-4o': { input: 2.50, output: 10.00 },
    'gpt-4-turbo': { input: 10.00, output: 30.00 },
    'o1-mini': { input: 3.00, output: 12.00 },
    'o1': { input: 15.00, output: 60.00 }
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
    
    // Translate all paragraphs at once
    const allText = paragraphsData.map(p => p.text).join('\n\n');
    const wordCount = allText.split(/\s+/).length;
    
    try {
        const result = await callChatGPT(allText, sourceLang, targetLang, apiKey, model);
        
        // Split translation back into paragraphs
        const translations = result.translation.split('\n\n').filter(t => t.trim());
        
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

// Call ChatGPT API
async function callChatGPT(text, fromLang, toLang, apiKey, model) {
    const targetLanguage = LANGUAGE_MAP[toLang];
    const mode = document.getElementById('translationMode').value;
    const style = document.getElementById('translationStyle').value;
    
    // Build system prompt based on mode and style
    let systemPrompt = buildSystemPrompt(fromLang, toLang, targetLanguage, mode, style);
    
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
