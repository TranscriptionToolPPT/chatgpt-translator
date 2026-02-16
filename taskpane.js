/* global Word, Office */

// تهيئة Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("✅ ChatGPT Translator جاهز للعمل!");
        
        // التحقق من وجود API Key محفوظ
        const savedKey = localStorage.getItem('openai_api_key');
        if (savedKey) {
            document.getElementById('apiKey').value = savedKey;
            document.getElementById('apiKeySaved').style.display = 'inline-block';
        }
    }
});

// حفظ الـ API Key
function saveApiKey() {
    const apiKey = document.getElementById('apiKey').value.trim();
    
    if (!apiKey) {
        showStatus('❌ من فضلك أدخل API Key', 'error');
        return;
    }
    
    if (!apiKey.startsWith('sk-')) {
        showStatus('❌ المفتاح غير صحيح. يجب أن يبدأ بـ sk-', 'error');
        return;
    }
    
    localStorage.setItem('openai_api_key', apiKey);
    document.getElementById('apiKeySaved').style.display = 'inline-block';
    showStatus('✅ تم حفظ المفتاح بنجاح', 'success');
}

// الترجمة الرئيسية
async function translateSelection() {
    const apiKey = localStorage.getItem('openai_api_key');
    
    if (!apiKey) {
        showStatus('❌ من فضلك أدخل وحفظ API Key أولاً', 'error');
        return;
    }
    
    showStatus('⏳ جاري الترجمة...', 'info');
    
    try {
        await Word.run(async (context) => {
            // الحصول على النص المحدد
            const range = context.document.getSelection();
            range.load('text, font');
            await context.sync();
            
            const selectedText = range.text;
            
            if (!selectedText || selectedText.trim().length === 0) {
                showStatus('❌ لم يتم تحديد أي نص. حدد النص المراد ترجمته.', 'error');
                return;
            }
            
            // الحصول على اللغات المختارة
            const sourceLang = document.getElementById('sourceLang').value;
            const targetLang = document.getElementById('targetLang').value;
            
            if (sourceLang === targetLang) {
                showStatus('⚠️ اللغة المصدر والهدف متطابقتان', 'error');
                return;
            }
            
            // استدعاء ChatGPT API
            const translation = await callChatGPT(selectedText, sourceLang, targetLang, apiKey);
            
            // استبدال النص بالترجمة (مع الحفاظ على التنسيق)
            range.insertText(translation, Word.InsertLocation.replace);
            
            await context.sync();
            
            showStatus('✅ تمت الترجمة بنجاح!', 'success');
        });
    } catch (error) {
        console.error('Translation error:', error);
        showStatus(`❌ حدث خطأ: ${error.message}`, 'error');
    }
}

// استدعاء ChatGPT API
async function callChatGPT(text, fromLang, toLang, apiKey) {
    // خريطة اللغات
    const langMap = {
        'es': 'Spanish',
        'en': 'English',
        'fr': 'French',
        'de': 'German',
        'it': 'Italian',
        'ar': 'Arabic'
    };
    
    const sourceLanguage = langMap[fromLang];
    const targetLanguage = langMap[toLang];
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: 'gpt-4o-mini', // أرخص وأسرع موديل
                messages: [
                    {
                        role: 'system',
                        content: `You are a professional translator. Translate the following text from ${sourceLanguage} to ${targetLanguage}. Return ONLY the translated text without any explanations, notes, or additional commentary.`
                    },
                    {
                        role: 'user',
                        content: text
                    }
                ],
                temperature: 0.3, // دقة أعلى
                max_tokens: 2000
            })
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error?.message || `HTTP Error: ${response.status}`);
        }
        
        const data = await response.json();
        
        if (!data.choices || !data.choices[0]) {
            throw new Error('استجابة غير متوقعة من API');
        }
        
        return data.choices[0].message.content.trim();
        
    } catch (error) {
        if (error.message.includes('fetch')) {
            throw new Error('فشل الاتصال بـ ChatGPT. تحقق من الإنترنت.');
        } else if (error.message.includes('Incorrect API key')) {
            throw new Error('API Key غير صحيح. تحقق من المفتاح.');
        } else if (error.message.includes('quota')) {
            throw new Error('نفذ رصيد API. أضف رصيد لحسابك على OpenAI.');
        }
        throw error;
    }
}

// عرض رسائل الحالة
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.style.display = 'block';
    
    // إخفاء الرسالة تلقائياً بعد 5 ثواني للرسائل الناجحة
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}
