// AI service: Integrates Gemini AI API via OpenRouter for FAQ chatbot assistance and proof report summarization.
const logger = require('../utils/logger');
const { primaryClientUrl } = require('../config/clientOrigins');

const apiKey = process.env.OPENROUTER_API_KEY;

if (apiKey) {
  logger.info('✅ OpenRouter AI initialized');
} else {
  logger.warn('⚠️  OPENROUTER_API_KEY not set. AI features will use mock responses.');
}

async function callOpenRouter(prompt) {
  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "HTTP-Referer": primaryClientUrl,
      "X-Title": "EcoSurvey",
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      "model": "google/gemini-2.5-flash",
      "max_tokens": 1000,
      "messages": [
        {"role": "user", "content": prompt}
      ]
    })
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`${response.status} ${errorText}`);
  }

  const data = await response.json();
  if (data && data.choices && data.choices.length > 0) {
    return data.choices[0].message.content;
  }
  throw new Error("Invalid response format from OpenRouter");
}

exports.answerFAQ = async (userQuestion, faqs) => {
  if (!apiKey) {
    return mockFAQAnswer(userQuestion, faqs);
  }

  try {
    const faqContext = faqs.map((f) => `Q: ${f.question}\nA: ${f.answer}`).join('\n\n');

    const prompt = `You are an AI Assistant for EcoSurvey.

Knowledge Base (FAQ):
---
${faqContext}
---

User's input: "${userQuestion}"

Instructions:
1. RESPONSE LANGUAGE (STRICT):
   - Match the EXACT language of the user's input (English input -> English response, Vietnamese input -> Vietnamese response).

2. DIRECT & CONCISE STYLE (CRITICAL):
   - DO NOT include conversational filler, fluff, or repetitive intros (e.g. DO NOT say "Hello! I am here to help...", "Regarding your question...", "Thank you for asking...").
   - Answer DIRECTLY to the point starting from the very first word.
   - Use clear formatting, bullet points (for multi-step instructions), and line breaks for readability.

3. GREETING INPUTS ONLY:
   - If the input is ONLY a simple greeting (e.g. "hi", "hello", "xin chào"): Reply with ONE concise welcome sentence (e.g. "Hello! How can I assist you with EcoSurvey today?" or "Xin chào! Bạn cần hỗ trợ thông tin gì về EcoSurvey?").

4. FAQ CONSTRAINTS & CROSS-LINGUAL MATCHING:
   - Cross-match semantic meaning between the user's input and the FAQ database (which may be in Vietnamese or English).
   - If the information is not in the FAQ database, state directly in one sentence that the information is not in the database and advise them to contact Admin.`;

    return await callOpenRouter(prompt);
  } catch (err) {
    logger.error('OpenRouter answerFAQ error:', err.message);
    return mockFAQAnswer(userQuestion, faqs);
  }
};

exports.summarizeReport = async (description, eventName) => {
  if (!apiKey) {
    return mockSummary(description, eventName);
  }

  try {
    const prompt = `Summarize the following environmental activity report in 2-3 concise sentences. Be objective and professional.

Event: ${eventName}
Report: ${description}

Provide only the summary, no introduction or extra text.`;

    return await callOpenRouter(prompt);
  } catch (err) {
    logger.error('OpenRouter summarizeReport error:', err.message);
    return mockSummary(description, eventName);
  }
};

// Mock data fallback function when AI API key is not configured.
function mockFAQAnswer(question, faqs) {
  const q = question.trim().toLowerCase();

  // Basic greeting check for mock fallback
  if (['hi', 'hello', 'hey', 'good morning', 'good afternoon', 'hi there'].includes(q)) {
    return "Hello! How can I assist you with EcoSurvey today?";
  }
  if (['chào', 'xin chào', 'chào bạn', 'alo', 'hi bạn'].includes(q)) {
    return "Xin chào! Tôi có thể giúp gì cho bạn về hệ thống EcoSurvey hôm nay?";
  }

  for (const faq of faqs) {
    const words = faq.question.toLowerCase().split(/\s+/);
    const matchCount = words.filter((w) => w.length > 3 && q.includes(w)).length;
    if (matchCount >= 2) return faq.answer;
  }

  // Detect basic English vs Vietnamese for fallback response
  const isEnglish = /[a-zA-Z]/.test(q) && !/[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/i.test(q);
  if (isEnglish) {
    return "I don't have information on this topic in the FAQ database. Please contact the Administrator for assistance.";
  }

  return "Tôi không có thông tin về chủ đề này. Vui lòng liên hệ Quản trị viên để được hỗ trợ chi tiết hơn.";
}

function mockSummary(description, eventName) {
  const sentences = description.replace(/\s+/g, ' ').trim().split(/[.!?]+/).filter(Boolean);
  const first2 = sentences.slice(0, 2).join('. ').trim();
  return `${first2}. Báo cáo hoạt động môi trường cho "${eventName}" thể hiện tinh thần tham gia tích cực các sáng kiến bền vững.`;
}
