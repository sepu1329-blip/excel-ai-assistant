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
    { id: 'p1', name: '표 형태로 정리하기', text: '바깥쪽과 안쪽 세로 실선 / 안쪽 가로선 점선 테두리 적용.' },
    { id: 'p2', name: '데이터 요약 분석하기', text: '현재 영역의 데이터를 분석하고 시사점 요약해줘.' },
    { id: 'p3', name: '번역', text: '선택한 영역의 값들을 한글로 번역해줘' },
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
    const updatePromptBtn = document.getElementById('update-prompt-btn');
    const cancelEditBtn = document.getElementById('cancel-edit-btn');
    const newPromptNameInput = document.getElementById('new-prompt-name');
    const newPromptTextInput = document.getElementById('new-prompt-text');
    const promptListEl = document.getElementById('prompt-list');

    // Additional Nav
    const homeBtn = document.getElementById('home-btn');

    let editingPromptId = null; // Track which prompt is being edited

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

    homeBtn.addEventListener('click', () => {
        if (!state.apiKey) {
            alert('먼저 API Key를 저장해주세요.');
            return;
        }
        showView('chat');
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
            // Add new
            const newPrompt = {
                id: 'p_' + Date.now(),
                name: name,
                text: text
            };
            state.savedPrompts.push(newPrompt);
            savePromptsAndRender();
        } else {
            alert('프롬프트 이름과 내용을 모두 입력해주세요.');
        }
    });

    updatePromptBtn.addEventListener('click', () => {
        if (!editingPromptId) return;
        const name = newPromptNameInput.value.trim();
        const text = newPromptTextInput.value.trim();

        if (name && text) {
            const pIndex = state.savedPrompts.findIndex(p => p.id === editingPromptId);
            if (pIndex > -1) {
                state.savedPrompts[pIndex].name = name;
                state.savedPrompts[pIndex].text = text;
            }
            savePromptsAndRender();
        } else {
            alert('프롬프트 이름과 내용을 모두 입력해주세요.');
        }
    });

    cancelEditBtn.addEventListener('click', resetPromptForm);

    function resetPromptForm() {
        editingPromptId = null;
        newPromptNameInput.value = '';
        newPromptTextInput.value = '';
        addPromptBtn.classList.remove('hidden');
        updatePromptBtn.classList.add('hidden');
        cancelEditBtn.classList.add('hidden');
        document.querySelectorAll('.prompt-item').forEach(el => el.classList.remove('editing'));
    }

    function savePromptsAndRender() {
        savePromptsToStorage();
        renderPrompts();
        resetPromptForm();
    }

    promptListEl.addEventListener('click', (e) => {
        const targetBtn = e.target.closest('.prompt-action-btn');
        const targetPromptInfo = e.target.closest('.prompt-item');

        // Handle Delete Button
        if (targetBtn && targetBtn.classList.contains('delete')) {
            const id = targetBtn.dataset.id;
            if (confirm('이 프롬프트를 삭제하시겠습니까?')) {
                state.savedPrompts = state.savedPrompts.filter(p => p.id !== id);
                if (editingPromptId === id) {
                    resetPromptForm();
                }
                savePromptsToStorage();
                renderPrompts();
            }
            return;
        }

        // Handle Click on Prompt Item (Edit)
        if (targetPromptInfo && !targetBtn) {
            const id = targetPromptInfo.dataset.id;
            const prompt = state.savedPrompts.find(p => p.id === id);
            if (prompt) {
                editingPromptId = id;
                newPromptNameInput.value = prompt.name;
                newPromptTextInput.value = prompt.text;

                // Toggle Buttons
                addPromptBtn.classList.add('hidden');
                updatePromptBtn.classList.remove('hidden');
                cancelEditBtn.classList.remove('hidden');

                // Highlight selected item
                document.querySelectorAll('.prompt-item').forEach(el => el.classList.remove('editing'));
                targetPromptInfo.classList.add('editing');
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
            div.dataset.id = p.id;
            if (p.id === editingPromptId) {
                div.classList.add('editing');
            }
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
                systemInstruction = `You are an Excel Agent. You must respond with EITHER a valid JSON array of actions to modify the spreadsheet OR a valid JSON object to ask for more information.

1. If the user's request is unclear or missing necessary details, you MUST ask a clarifying question.
   - Example missing details: asked to remove duplicates but no range specified AND you can't infer it from Context. 
   - Note: If no range is specified by the user, you SHOULD default to using the 'Range' provided in the Context.
   - Request format: {"action": "ask_question", "message": "중복을 제거할 데이터 범위가 어디인가요? (예: A1:D10)"}

2. If you have enough information, output a JSON array of actions.
Supported actions:
- {"action": "set_values", "range": "A1:B2", "values": [["1", "2"], ["3", "4"]]}
- {"action": "format_color", "range": "A1:A5", "color": "#FF0000"} // Hex color
- {"action": "format_borders", "range": "A1:B2", "style": "Continuous", "weight": "Thin"} // Available styles: None, Continuous, Dash. Weights: Hairline, Thin, Medium, Thick
- {"action": "clear_range", "range": "A1:Z100"}
- {"action": "set_formula", "range": "C1", "formula": "=A1+B1"}
- {"action": "add_chart", "type": "ColumnClustered", "range": "A1:B5", "title": "My Chart"} // Available types: ColumnClustered, Line, Pie
- {"action": "modify_chart", "chart_name": "Chart 1", "title": "New Title", "series_colors": ["#FF0000"]}
- {"action": "set_row_height", "range": "A1:A5", "height": 50}
- {"action": "set_column_width", "range": "A1:C1", "width": 100}
- {"action": "hide_rows", "range": "A1:A5", "hidden": true}
- {"action": "hide_columns", "range": "A1:C1", "hidden": true}
- {"action": "remove_duplicates", "range": "A1:C10", "columns": [0, 1], "includesHeader": true} // columns is optional array of 0-based column indices to check for duplicates. includesHeader is optional boolean.

DO NOT wrap the JSON in markdown code blocks like \`\`\`json. Just output the raw JSON array (or object for questions) directly. If you cannot fulfill the request, output an empty array [].`;
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
                    let parsedJson = JSON.parse(cleanJsonStr);

                    // Normalize to an array
                    const actionsArray = Array.isArray(parsedJson) ? parsedJson : [parsedJson];

                    const questionAction = actionsArray.find(a => a.action === 'ask_question');

                    if (questionAction) {
                        // AI is asking for clarification
                        appendMessage('ai', `<p>🙋 <b>질문:</b> ${questionAction.message}</p>`);
                    } else if (actionsArray.some(a => a.action)) {
                        await executeExcelActions(actionsArray);
                        appendMessage('ai', `<p>✅ Executed ${actionsArray.length} action(s) successfully.</p>`);
                    } else {
                        // Let the user know the AI couldn't generate valid actions, print raw output for debug
                        appendMessage('ai', `<p>No valid actions found to apply.</p><div style="font-size: 0.85em; color: gray; margin-top: 5px; padding: 5px; background: #f9f9f9; border-radius: 4px; overflow-x: auto;"><b>Raw AI Output:</b><br/>${responseText.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div>`);
                    }
                } catch (err) {
                    console.error("Agent JSON parsing error:", err, responseText);
                    appendMessage('error', `<p>Failed to parse AI response as valid format.<br/><br/><div style="font-size: 0.85em; color: gray;"><b>Raw AI Output:</b><br/>${responseText.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div></p>`);
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
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            let range;
            if (contextType === 'cell') {
                range = context.workbook.getSelectedRange();
            } else {
                range = sheet.getUsedRange();
            }
            range.load("values, address");

            // Load charts
            const charts = sheet.charts;
            charts.load("items/name, items/title/text, items/type");

            await context.sync();

            let chartsInfo = charts.items.map(c => ({
                name: c.name,
                title: c.title ? c.title.text : "",
                type: c.type
            }));

            let contextStr = `Range: ${range.address}\nValues: ${JSON.stringify(range.values)}\nCharts: ${JSON.stringify(chartsInfo)}`;
            resolve(contextStr);
        }).catch(reject);
    });
}

async function executeExcelActions(actions) {
    if (!window.Excel || !window.Excel.run) {
        console.warn("Not running in Excel, mocked execution of:", actions);
        return;
    }

    return Excel.run({ mergeUndoGroup: true }, async (context) => {
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
                    const borderTypes = act.border_types || ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];

                    borderTypes.forEach(b => {
                        try {
                            range.format.borders.getItem(b).style = style;
                            range.format.borders.getItem(b).weight = weight;
                        } catch (err) {
                            console.warn("Unsupported border type or error applying type:", b, err);
                        }
                    });
                }
                else if (act.action === 'clear_range') {
                    range.clear();
                }
                else if (act.action === 'set_formula' && act.formula) {
                    range.formulas = [[act.formula]];
                }
                else if (act.action === 'set_row_height' && act.height !== undefined) {
                    range.format.rowHeight = act.height;
                }
                else if (act.action === 'set_column_width' && act.width !== undefined) {
                    range.format.columnWidth = act.width;
                }
                else if (act.action === 'hide_rows' && act.hidden !== undefined) {
                    range.rowHidden = act.hidden;
                }
                else if (act.action === 'hide_columns' && act.hidden !== undefined) {
                    range.columnHidden = act.hidden;
                }
                else if (act.action === 'remove_duplicates') {
                    // columns is an array of 0-based column indices relative to the range, e.g., [0, 1] means check 1st and 2nd col of the range
                    // includesHeader is a boolean
                    let columns = act.columns;
                    const includesHeader = act.includesHeader !== undefined ? act.includesHeader : true;

                    if (!columns || columns.length === 0) {
                        // If no specific columns provided, default to all columns in the range
                        range.load("columnCount");
                        await context.sync();

                        columns = Array.from({ length: range.columnCount }, (_, i) => i);
                    }

                    // removeDuplicates takes (columns: number[], includesHeader: boolean)
                    range.removeDuplicates(columns, includesHeader);
                }
                else if (act.action === 'add_chart') {
                    const chartType = act.type || "ColumnClustered";
                    const chart = sheet.charts.add(chartType, range, "Auto");
                    if (act.title) {
                        chart.title.text = act.title;
                    }
                }
                else if (act.action === 'modify_chart' && act.chart_name) {
                    const chart = sheet.charts.getItem(act.chart_name);

                    if (act.title) {
                        chart.title.text = act.title;
                    }

                    // Handle Series Colors Update
                    if (act.series_colors && Array.isArray(act.series_colors)) {
                        const seriesCollection = chart.series;
                        seriesCollection.load("count");
                        await context.sync(); // Sync to get the count

                        for (let i = 0; i < act.series_colors.length; i++) {
                            if (i < seriesCollection.count) {
                                const seriesItem = seriesCollection.getItemAt(i);
                                seriesItem.format.fill.setSolidColor(act.series_colors[i]);
                            }
                        }
                    }
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
