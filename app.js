// Wait for Office.js to initialize
Office.onReady((info) => {
    if (info.host === Office.HostType.Workbook) {
        console.log("Office.js is ready in Excel.");
    }
    initApp();
});

// App State
const state = {
    apiKey: '',
    model: 'gemini-2.5-flash',
    mode: 'ask', // 'ask' or 'agent'
    context: 'cell', // 'cell' or 'sheet'
    chatHistory: [], // store messages
    savedPrompts: [] // array of { id, name, text }
};

const defaultPrompts = [
    { id: 'p1', name: '표 형태로 정리하기', text: '선택한 영역의 데이터들을 표 형태로 깔끔하게 정리해줘.' },
    { id: 'p2', name: '데이터 요약 분석하기', text: '현재 영역의 데이터들을 분석하고 시사점 요약해줘.' },
    { id: 'p3', name: '영역 바깥쪽 실선 테두리', text: '선택한 영역의 바깥쪽 실선 테두리를 적용해줘.' },
    { id: 'p4', name: '빈 셀 노란색 칠하기', text: '값이 비어있는 셀을 찾아서 노란색으로 칠해줘.' }
];

function initApp() {
    // DOM Elements
    const settingsView = document.getElementById('settings-view');
    const chatView = document.getElementById('chat-view');
    const apiKeyInput = document.getElementById('api-key');
    const modelSelect = document.getElementById('model-select');
    const saveSettingsBtn = document.getElementById('save-settings-btn');

    const settingsBtn = document.getElementById('settings-btn');
    const clearChatBtn = document.getElementById('clear-chat-btn');
    const chatHistoryEl = document.getElementById('chat-history');
    const chatInput = document.getElementById('chat-input');
    const sendBtn = document.getElementById('send-btn');
    const savedPromptsSelect = document.getElementById('saved-prompts-select');

    // Prompt Management Elements
    const addPromptBtn = document.getElementById('add-prompt-btn');
    const newPromptNameInput = document.getElementById('new-prompt-name');
    const newPromptTextInput = document.getElementById('new-prompt-text');
    const promptListEl = document.getElementById('prompt-list');

    // Toggles
    const modeBtns = document.querySelectorAll('.mode-toggle .toggle-btn');
    const contextBtns = document.querySelectorAll('.context-toggle .toggle-btn');

    // Load saved settings
    const savedKey = localStorage.getItem('gemini_api_key');
    const savedModel = localStorage.getItem('gemini_model');
    const savedPromptsLocal = localStorage.getItem('gemini_prompts');

    // Initialize Prompts
    if (savedPromptsLocal) {
        try {
            state.savedPrompts = JSON.parse(savedPromptsLocal);
        } catch (e) {
            state.savedPrompts = [...defaultPrompts];
        }
    } else {
        state.savedPrompts = [...defaultPrompts];
        savePromptsToStorage();
    }
    renderPrompts();

    if (savedKey) {
        state.apiKey = savedKey;
        apiKeyInput.value = savedKey;
        if (savedModel) {
            state.model = savedModel;
            modelSelect.value = savedModel;
        }
        showView('chat');
    }

    // Event Listeners
    saveSettingsBtn.addEventListener('click', () => {
        const key = apiKeyInput.value.trim();
        if (key) {
            state.apiKey = key;
            state.model = modelSelect.value;
            localStorage.setItem('gemini_api_key', key);
            localStorage.setItem('gemini_model', state.model);
            showView('chat');
        } else {
            alert('Please enter a valid API key.');
        }
    });

    settingsBtn.addEventListener('click', () => {
        showView('settings');
    });

    clearChatBtn.addEventListener('click', () => {
        state.chatHistory = [];
        chatHistoryEl.innerHTML = `
            <div class="message system">
              <p>Chat cleared. Welcome to Gemini AI!</p>
            </div>
        `;
    });

    // Toggle logic
    modeBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            modeBtns.forEach(b => b.classList.remove('active'));
            const target = e.target;
            target.classList.add('active');
            state.mode = target.dataset.mode;
        });
    });

    contextBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            contextBtns.forEach(b => b.classList.remove('active'));
            const target = e.target;
            target.classList.add('active');
            state.context = target.dataset.context;
        });
    });

    // Prompt Management Logic
    addPromptBtn.addEventListener('click', () => {
        const name = newPromptNameInput.value.trim();
        const text = newPromptTextInput.value.trim();

        if (name && text) {
            const newPrompt = {
                id: 'p_' + Date.now(),
                name: name,
                text: text
            };
            state.savedPrompts.push(newPrompt);
            savePromptsToStorage();
            renderPrompts();
            newPromptNameInput.value = '';
            newPromptTextInput.value = '';
        } else {
            alert('프롬프트 이름과 내용을 모두 입력해주세요.');
        }
    });

    promptListEl.addEventListener('click', (e) => {
        const target = e.target.closest('.prompt-action-btn');
        if (!target) return;

        const id = target.dataset.id;
        if (target.classList.contains('delete')) {
            if (confirm('이 프롬프트를 삭제하시겠습니까?')) {
                state.savedPrompts = state.savedPrompts.filter(p => p.id !== id);
                savePromptsToStorage();
                renderPrompts();
            }
        }
    });

    function savePromptsToStorage() {
        localStorage.setItem('gemini_prompts', JSON.stringify(state.savedPrompts));
    }

    function renderPrompts() {
        // Render Settings List
        promptListEl.innerHTML = '';
        state.savedPrompts.forEach(p => {
            const div = document.createElement('div');
            div.className = 'prompt-item';
            div.innerHTML = `
                <div class="prompt-info">
                    <span class="prompt-name">${p.name}</span>
                    <span class="prompt-text">${p.text}</span>
                </div>
                <div class="prompt-actions">
                    <button class="prompt-action-btn delete" data-id="${p.id}" title="삭제">❌</button>
                </div>
            `;
            promptListEl.appendChild(div);
        });

        // Render Dropdown options
        savedPromptsSelect.innerHTML = '<option value="">-- 자주 쓰는 프롬프트 선택 --</option>';
        state.savedPrompts.forEach(p => {
            const option = document.createElement('option');
            option.value = p.text;
            option.textContent = p.name;
            savedPromptsSelect.appendChild(option);
        });
    }

    // Saved Prompts logic (Dropdown Change)
    savedPromptsSelect.addEventListener('change', (e) => {
        const text = e.target.value;
        if (text) {
            chatInput.value = text;
        }
        // Reset dropdown to default after selection
        savedPromptsSelect.value = "";
    });

    // Chat input handling
    chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    });

    sendBtn.addEventListener('click', handleSend);

    function showView(viewName) {
        if (viewName === 'settings') {
            settingsView.classList.remove('hidden');
            chatView.classList.add('hidden');
        } else {
            settingsView.classList.add('hidden');
            chatView.classList.remove('hidden');
        }
    }

    async function handleSend() {
        const text = chatInput.value.trim();
        if (!text) return;
        if (!state.apiKey) {
            showView('settings');
            return;
        }

        // Add user message to UI
        appendMessage('user', text);
        chatInput.value = '';

        // Disable input while loading
        chatInput.disabled = true;
        sendBtn.disabled = true;

        // Show loading
        const loadingId = appendLoading();

        try {
            // 1. Gather context from Excel
            const excelContextData = await getExcelContext(state.context);

            // 2. Call Gemini
            let systemInstruction = "";
            let promptText = "";

            if (state.mode === 'ask') {
                systemInstruction = "You are a helpful Excel assistant. Answer questions based on the provided context.";
                promptText = `Context:\n${excelContextData}\n\nUser Question:\n${text}`;
            } else { // Agent mode
                systemInstruction = `You are an Excel Agent. You can ONLY respond with a valid JSON array of actions to modify the spreadsheet. 
Supported actions:
- {"action": "set_values", "range": "A1:B2", "values": [["1", "2"], ["3", "4"]]}
- {"action": "format_color", "range": "A1:A5", "color": "#FF0000"} // Hex color
- {"action": "format_borders", "range": "A1:B2", "style": "Continuous", "weight": "Thin"} // Available styles: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot. Weights: Hairline, Thin, Medium, Thick
- {"action": "clear_range", "range": "A1:Z100"}
- {"action": "set_formula", "range": "C1", "formula": "=A1+B1"}
DO NOT wrap the JSON in markdown code blocks like \`\`\`json. Just output the raw JSON array. If you cannot fulfill the request, output an empty array [].`;
                promptText = `Context:\n${excelContextData}\n\nUser Request:\n${text}`;
            }

            const responseText = await callGeminiAPI(systemInstruction, promptText);

            // Remove loading
            document.getElementById(loadingId).remove();

            if (state.mode === 'ask') {
                appendMessage('ai', marked.parse(responseText));
            } else {
                // Agent mode: Attempt to parse JSON and execute
                try {
                    // Try to clean markdown block if the AI ignored instructions
                    let cleanJsonStr = responseText.replace(/```json/gi, '').replace(/```/g, '').trim();
                    const actions = JSON.parse(cleanJsonStr);

                    if (Array.isArray(actions) && actions.length > 0) {
                        await executeExcelActions(actions);
                        appendMessage('ai', `<p>✅ Executed ${actions.length} action(s) successfully.</p>`);
                    } else {
                        appendMessage('ai', `<p>No valid actions found to apply.</p>`);
                    }
                } catch (err) {
                    console.error("Agent JSON parsing error:", err, responseText);
                    appendMessage('error', `<p>Failed to parse AI response as valid actions. View console for details.</p>`);
                }
            }

        } catch (error) {
            console.error(error);
            document.getElementById(loadingId)?.remove();
            appendMessage('error', `<p>Error: ${error.message}</p>`);
        } finally {
            chatInput.disabled = false;
            sendBtn.disabled = false;
            chatInput.focus();
        }
    }

    function appendMessage(role, contentMarkup) {
        const div = document.createElement('div');
        div.className = `message ${role}`;

        if (role === 'user') {
            div.textContent = contentMarkup; // User input is raw text
        } else {
            div.innerHTML = contentMarkup; // AI is parsed markdown / HTML
        }

        chatHistoryEl.appendChild(div);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
    }

    function appendLoading() {
        const id = 'loading-' + Date.now();
        const div = document.createElement('div');
        div.id = id;
        div.className = 'loading-indicator';
        div.innerHTML = `<div class="dot"></div><div class="dot"></div><div class="dot"></div>`;
        chatHistoryEl.appendChild(div);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
        return id;
    }
}

// ============== EXCEL LOGIC ==============

async function getExcelContext(contextType) {
    return new Promise((resolve, reject) => {
        // Check if Office is initialized (will fail outside Excel if not properly mocked)
        if (!window.Excel || !window.Excel.run) {
            console.warn("Not running in Excel, mocking context.");
            return resolve("Range: A1\nValues: [[\"Mock\", \"Data\"]]");
        }

        Excel.run(async (context) => {
            let range;
            if (contextType === 'cell') {
                range = context.workbook.getSelectedRange();
            } else {
                range = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            }
            range.load("values, address");
            await context.sync();

            let contextStr = `Range: ${range.address}\nValues: ${JSON.stringify(range.values)}`;
            resolve(contextStr);
        }).catch(reject);
    });
}

async function executeExcelActions(actions) {
    if (!window.Excel || !window.Excel.run) {
        console.warn("Not running in Excel, mocked execution of:", actions);
        return;
    }

    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        for (const act of actions) {
            try {
                if (!act.range) continue;
                const range = sheet.getRange(act.range);

                if (act.action === 'set_values' && act.values) {
                    range.values = act.values;
                }
                else if (act.action === 'format_color' && act.color) {
                    range.format.fill.color = act.color;
                }
                else if (act.action === 'format_borders') {
                    const style = act.style || "Continuous";
                    const weight = act.weight || "Thin";
                    range.format.borders.getItem('EdgeTop').style = style;
                    range.format.borders.getItem('EdgeTop').weight = weight;
                    range.format.borders.getItem('EdgeBottom').style = style;
                    range.format.borders.getItem('EdgeBottom').weight = weight;
                    range.format.borders.getItem('EdgeLeft').style = style;
                    range.format.borders.getItem('EdgeLeft').weight = weight;
                    range.format.borders.getItem('EdgeRight').style = style;
                    range.format.borders.getItem('EdgeRight').weight = weight;
                }
                else if (act.action === 'clear_range') {
                    range.clear();
                }
                else if (act.action === 'set_formula' && act.formula) {
                    range.formulas = [[act.formula]];
                }
            } catch (err) {
                console.error("Error executing action:", act, err);
            }
        }
        await context.sync();
    });
}

// ============== GEMINI API LOGIC ==============

async function callGeminiAPI(systemInstruction, userPrompt) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${state.model}:generateContent?key=${state.apiKey}`;

    const body = {
        contents: [{
            role: "user",
            parts: [{ text: userPrompt }]
        }],
        systemInstruction: {
            parts: [{ text: systemInstruction }]
        },
        generationConfig: {
            temperature: 0.1
        }
    };

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || "Failed to fetch from Gemini API");
    }

    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}
