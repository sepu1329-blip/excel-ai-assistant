Office.onReady((info) => {
    console.log("Office.js is ready.");
    initApp();
});

// Fallback initialization if Office.onReady is slow (wait 2 seconds)
setTimeout(() => {
    if (typeof window.appInitialized === 'undefined') {
        console.log("Fallback initialization acting...");
        initApp();
    }
}, 2000);

// App State
const state = {
    apiKey: '',
    model: 'gemini-3.1-flash',
    mode: 'ask', // 'ask' or 'agent'
    context: 'cell', // 'cell' or 'sheet'
    chatHistory: [], // store messages
    savedPrompts: [] // array of { id, name, text }
};

let currentAbortController = null; // To cancel AI requests

const defaultPrompts = [
    { id: 'p1', name: '표 형태로 정리하기', text: '바깥쪽과 안쪽 세로 실선 / 안쪽 가로선 점선 테두리 적용.' },
    { id: 'p2', name: '데이터 요약 분석하기', text: '현재 영역의 데이터를 분석하고 시사점 요약해줘.' },
    { id: 'p3', name: '번역', text: '선택한 영역의 값들을 한글로 번역해줘' },
    { id: 'p4', name: '빈 셀 노란색 칠하기', text: '값이 비어있는 셀을 찾아서 노란색으로 칠해줘.' },
    { id: 'p5', name: '천 단위 구분 기호', text: '선택한 숫자에 천 단위 구분 기호(,)를 적용해줘.' },
    { id: 'p6', name: '날짜 형식 (YYYY-MM-DD)', text: '날짜를 YYYY-MM-DD 형식으로 바꿔줘.' }
];

function initApp() {
    if (window.appInitialized) return;
    window.appInitialized = true;
    console.log("initApp started");
    // DOM Elements
    const settingsView = document.getElementById('settings-view');
    const chatView = document.getElementById('chat-view');
    const apiKeyInput = document.getElementById('api-key');
    const modelSelect = document.getElementById('model-select');
    const saveSettingsBtn = document.getElementById('save-settings-btn');
    if (!saveSettingsBtn) {
        alert("CRITICAL ERROR: 'save-settings-btn' not found in DOM.");
        return;
    }

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

    // Load saved settings (Wrapped in try-catch for restricted environments)
    let savedKey = null;
    let savedModel = null;
    let savedPromptsLocal = null;
    
    try {
        savedKey = localStorage.getItem('gemini_api_key');
        savedModel = localStorage.getItem('gemini_model');
        savedPromptsLocal = localStorage.getItem('gemini_prompts');
    } catch (e) {
        console.warn("Storage access denied:", e);
    }

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
        console.log("Save button clicked");
        const key = apiKeyInput.value.trim();
        if (key) {
            state.apiKey = key;
            state.model = modelSelect.value;
            try {
                localStorage.setItem('gemini_api_key', key);
                localStorage.setItem('gemini_model', state.model);
            } catch (e) {
                console.warn("Could not save to localStorage:", e);
            }
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
        try {
            localStorage.setItem('gemini_prompts', JSON.stringify(state.savedPrompts));
        } catch (e) {
            console.warn("Could not save prompts:", e);
        }
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
            settingsView.classList.add('visible');
            chatView.classList.add('hidden');
            chatView.classList.remove('visible');
        } else {
            settingsView.classList.add('hidden');
            settingsView.classList.remove('visible');
            chatView.classList.remove('hidden');
            chatView.classList.add('visible');
        }
    }

    async function handleSend() {
        // If already loading, this acts as a STOP button
        if (currentAbortController) {
            currentAbortController.abort();
            currentAbortController = null;
            return;
        }

        const text = chatInput.value.trim();
        if (!text) return;
        if (!state.apiKey) {
            showView('settings');
            return;
        }

        // Add user message to UI
        appendMessage('user', text);
        chatInput.value = '';

        // Disable input
        chatInput.disabled = true;
        
        // Change button to STOP mode
        sendBtn.classList.add('loading');
        sendBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="6" y="6" width="12" height="12"></rect>
            </svg>
        `;

        // Create new AbortController
        currentAbortController = new AbortController();

        // Show loading
        const loadingId = appendLoading("엑셀 데이터를 읽고 있습니다...");

        try {
            // 1. Gather context from Excel
            const excelContextData = await getExcelContext(state.context);
            
            updateLoadingStatus(loadingId, "Gemini AI가 생각 중입니다...");

            // 2. Call Gemini
            let systemInstruction = "";
            let promptText = "";

            if (state.mode === 'ask') {
                systemInstruction = "You are a helpful Excel assistant. Answer questions based on the provided context.";
                promptText = `Context:\n${excelContextData}\n\nUser Question:\n${text}`;
            } else { // Agent mode
                systemInstruction = `You are an Excel Agent. You must respond with EITHER a valid JSON array of actions to modify the spreadsheet OR a valid JSON object to ask for more information.

1. CRITICAL: For ANY action that requires a "range", if the user does NOT explicitly specify a range (like A1:B10) in their prompt, you MUST use the "Range" value provided in the Context. DO NOT ask the user for a range if one is provided in the Context.

2. If the user's request is unclear or missing necessary details (other than range), you MUST ask a clarifying question.
   - Example 1: asked to remove duplicates, but didn't specify whether to just delete the copies (keeping the first occurrence) or to delete ALL duplicates including the original. ALWAYS ask this question for remove duplicates if not specified.
   - Example 2: asked to apply a format (like strikethrough, bold) to a specific line inside a multiline cell (e.g. "strike through the second line"). Office.js CANNOT format partial text natively. You MUST ask the user which line to modify if not specified, and use the 'modify_cell_lines' action to prepend text like '[취소] ' to that specific line.
   - Request format: {"action": "ask_question", "message": "몇 번째 줄에 취소선(또는 강조) 표시를 추가할까요?"}

3. If you have enough information, output a JSON array of actions.
Supported actions:
- {"action": "set_values", "range": "A1:B2", "values": [["1", "2"], ["3", "4"]]}
- {"action": "format_color", "range": "A1:A5", "color": "#FF0000"}
- {"action": "format_borders", "range": "A1:B2", "style": "Continuous", "weight": "Thin"}
- {"action": "format_font", "range": "A1:A5", "strikethrough": true, "size": 14, "bold": true}
- {"action": "format_alignment", "range": "A1:A5", "horizontal": "Center", "vertical": "Center"} // horizontal: Left|Center|Right. vertical: Top|Center|Bottom
- {"action": "clear_range", "range": "A1:Z100"}
- {"action": "set_formula", "range": "C1", "formula": "=A1+B1"}
- {"action": "add_chart", "type": "ColumnClustered", "range": "A1:B5", "title": "My Chart"} // type: ColumnClustered, Line, Pie, BarClustered
- {"action": "add_table", "range": "A1:C5", "has_header": true}
- {"action": "insert_cells", "range": "A2:A2", "shift": "Down"} // shift: "Down" or "Right"
- {"action": "delete_cells", "range": "A2:A2", "shift": "Up"} // shift: "Up" or "Left"
- {"action": "modify_chart", "chart_name": "Chart 1", "title": "New Title", "series_colors": ["#FF0000"]}
- {"action": "set_row_height", "range": "A1:A5", "height": 50}
- {"action": "set_column_width", "range": "A1:C1", "width": 100}
- {"action": "hide_rows", "range": "A1:A5", "hidden": true}
- {"action": "hide_columns", "range": "A1:C1", "hidden": true}
- {"action": "remove_duplicates", "range": "A1:C10", "columns": [0, 1], "includesHeader": false, "keepFirst": true}
- {"action": "apply_filter", "range": "A1:C10", "column": 0, "criterion1": ">=3"} // custom filter on 0-based column. Uses operators like >, <, =, >=.
- {"action": "apply_filter", "range": "A1:C10"} // Just adding a filter dropdown to the top row without specific criteria. Do NOT ask for a column if the user just asks to "add a filter".
- {"action": "clear_filter"} // clears all filters on the sheet
- {"action": "set_print_area", "range": "A1:D20"}
- {"action": "merge_identical_cells", "range": "A1:A10", "direction": "vertical"} // merges adjacent cells with identical values. direction: "vertical" or "horizontal"
- {"action": "add_data_validation", "range": "A1:A10", "source": "1,2,3"} // Creates an in-cell dropdown with the comma-separated list or Excel range formula
- {"action": "modify_cell_lines", "range": "A1:A2", "line_index": 1, "prefix": "[취소] ", "suffix": ""} // 0-based index of the line to modify in a multiline cell. A workaround for partial formatting constraints.
- {"action": "set_number_format", "range": "A1:A10", "format": "0.00%"} // Applies a custom number format string (e.g., "#,##0", "yyyy-mm-dd", "0.00%", etc.)

DO NOT wrap the JSON in markdown code blocks like \`\`\`json. Just output the raw JSON array (or object for questions) directly. If you cannot fulfill the request, output an empty array [].`;
                promptText = `Context:\n${excelContextData}\n\nUser Request:\n${text}`;
            }

            const responseText = await callGeminiAPI(systemInstruction, promptText, currentAbortController.signal);

            // Remove loading
            document.getElementById(loadingId)?.remove();

            if (state.mode === 'ask') {
                appendMessage('ai', marked.parse(responseText));
            } else {
                updateLoadingStatus(loadingId, "AI 응답을 분석하고 엑셀 작업을 준비 중입니다...");
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
                        updateLoadingStatus(loadingId, "엑셀 작업을 실행하고 있습니다...");
                        await executeExcelActions(actionsArray);
                        document.getElementById(loadingId)?.remove();
                        appendMessage('ai', `<p>✅ Executed ${actionsArray.length} action(s) successfully.</p>`);
                    } else {
                        document.getElementById(loadingId)?.remove();
                        // Let the user know the AI couldn't generate valid actions, print raw output for debug
                        appendMessage('ai', `<p>No valid actions found to apply.</p><div style="font-size: 0.85em; color: gray; margin-top: 5px; padding: 5px; background: #f9f9f9; border-radius: 4px; overflow-x: auto;"><b>Raw AI Output:</b><br/>${responseText.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div>`);
                    }
                } catch (err) {
                    console.error("Agent JSON parsing error:", err, responseText);
                    document.getElementById(loadingId)?.remove();
                    appendMessage('error', `<p>Failed to parse AI response as valid format.<br/><br/><div style="font-size: 0.85em; color: gray;"><b>Raw AI Output:</b><br/>${responseText.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div></p>`);
                }
            }

        } catch (error) {
            if (error.name === 'AbortError') {
                console.log("AI request aborted by user.");
                document.getElementById(loadingId)?.remove();
                appendMessage('system', "작업이 사용자에 의해 중단되었습니다.");
            } else {
                console.error(error);
                document.getElementById(loadingId)?.remove();
                appendMessage('error', `<p>Error: ${error.message}</p>`);
            }
        } finally {
            currentAbortController = null;
            chatInput.disabled = false;
            
            // Restore button to SEND mode
            sendBtn.classList.remove('loading');
            sendBtn.innerHTML = `
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="22" y1="2" x2="11" y2="13"></line>
                    <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
                </svg>
            `;
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

    function appendLoading(initialMessage = "AI가 생각 중입니다...") {
        const id = 'loading-' + Date.now();
        const div = document.createElement('div');
        div.id = id;
        div.className = 'loading-indicator';
        div.innerHTML = `
            <div class="loading-dots">
                <div class="dot"></div>
                <div class="dot"></div>
                <div class="dot"></div>
            </div>
            <div class="status-text">${initialMessage}</div>
        `;
        chatHistoryEl.appendChild(div);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
        return id;
    }

    function updateLoadingStatus(id, message) {
        const el = document.getElementById(id);
        if (el) {
            const statusEl = el.querySelector('.status-text');
            if (statusEl) {
                statusEl.textContent = message;
            }
        }
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
            // const charts = sheet.charts;
            // charts.load("items/name, items/title/text, items/type");

            await context.sync();

            // let chartsInfo = charts.items.map(c => ({
            //     name: c.name,
            //     title: c.title ? c.title.text : "",
            //     type: c.type
            // }));

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

    return Excel.run({ delayForCellEdit: true }, async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        for (const act of actions) {
            try {
                if (act.action === 'clear_filter') {
                    sheet.autoFilter.clearCriteria();
                    continue;
                }

                if (!act.range) continue;
                const range = sheet.getRange(act.range);

                if (act.action === 'set_values' && act.values) {
                    range.values = act.values;
                }
                else if (act.action === 'format_color' && act.color) {
                    range.format.fill.color = act.color;
                }
                else if (act.action === 'format_font') {
                    if (act.strikethrough !== undefined) range.format.font.strikethrough = act.strikethrough;
                    if (act.size !== undefined) range.format.font.size = act.size;
                    if (act.bold !== undefined) range.format.font.bold = act.bold;
                    if (act.italic !== undefined) range.format.font.italic = act.italic;
                }
                else if (act.action === 'format_alignment') {
                    if (act.horizontal) range.format.horizontalAlignment = act.horizontal;
                    if (act.vertical) range.format.verticalAlignment = act.vertical;
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
                    // Office.js 'removeDuplicates' expects (columns: number[], includesHeader: boolean)
                    // The 'columns' array should contain the 0-indexed column numbers to use for identifying duplicates.
                    // By default, Excel's removeDuplicates keeps the *first* occurrence.
                    // If the user wants to delete ALL occurrences (keepFirst: false), this requires custom logic.

                    let columns = act.columns;
                    const includesHeader = act.includesHeader === true;
                    const keepFirst = act.keepFirst !== false; // defaults to true

                    if (!columns || columns.length === 0) {
                        range.load("columnCount");
                        await context.sync();
                        columns = Array.from({ length: range.columnCount }, (_, i) => i);
                    }

                    if (keepFirst) {
                        // Native Excel behavior
                        const result = range.removeDuplicates(columns, includesHeader);
                        result.load("removed");
                    } else {
                        // Delete ALL duplicates (including the original)
                        range.load("values, rowCount");
                        await context.sync();

                        const values = range.values;
                        const startIndex = includesHeader ? 1 : 0;
                        const rowKeys = values.map(row => columns.map(c => row[c]).join('|'));

                        // Count frequencies
                        const freq = {};
                        for (let i = startIndex; i < rowKeys.length; i++) {
                            freq[rowKeys[i]] = (freq[rowKeys[i]] || 0) + 1;
                        }

                        // Collect rows to clear
                        // We iterate backwards to avoid shifting row indices when performing operations
                        // (Though we are using clear() + filter, a more robust way is clearing duplicate rows)
                        let clearedCount = 0;
                        for (let i = rowKeys.length - 1; i >= startIndex; i--) {
                            if (freq[rowKeys[i]] > 1) {
                                // Clear this row within the range
                                const rowRange = range.getRow(i);
                                rowRange.clear();
                                clearedCount++;
                            }
                        }
                        console.log(`Cleared ${clearedCount} rows of complete duplicates.`);
                    }
                }
                else if (act.action === 'add_chart') {
                    const chartType = act.type || "ColumnClustered";
                    const chart = sheet.charts.add(chartType, range, "Auto");
                    if (act.title) {
                        chart.title.text = act.title;
                        chart.title.visible = true;
                    }
                }
                else if (act.action === 'add_table') {
                    const hasHeader = act.has_header !== false;
                    const newTable = sheet.tables.add(range, hasHeader);
                    if (act.table_name) {
                        newTable.name = act.table_name;
                    }
                }
                else if (act.action === 'insert_cells') {
                    range.insert(act.shift === "Right" ? "Right" : "Down");
                }
                else if (act.action === 'delete_cells') {
                    range.delete(act.shift === "Left" ? "Left" : "Up");
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
                else if (act.action === 'add_data_validation') {
                    // source should be a comma separated string like "1,2,3"
                    range.dataValidation.rule = {
                        list: {
                            inCellDropDown: true,
                            source: act.source
                        }
                    };
                }
                else if (act.action === 'set_print_area') {
                    sheet.pageLayout.setPrintArea(range);
                }
                else if (act.action === 'apply_filter') {
                    if (act.criterion1 !== undefined && act.column !== undefined) {
                        sheet.autoFilter.apply(range, act.column, {
                            filterOn: Excel.FilterOn.custom,
                            criterion1: act.criterion1
                        });
                    } else {
                        // Just apply the filter dropdowns to the range without specific criteria
                        sheet.autoFilter.apply(range);
                    }
                }
                else if (act.action === 'merge_identical_cells') {
                    const direction = act.direction || "vertical";
                    range.load("values, rowCount, columnCount");
                    await context.sync();

                    const values = range.values;
                    const rows = range.rowCount;
                    const cols = range.columnCount;

                    if (direction === "vertical") {
                        for (let c = 0; c < cols; c++) {
                            let startR = 0;
                            while (startR < rows) {
                                let endR = startR;
                                while (endR + 1 < rows && values[endR][c] === values[endR + 1][c] && values[endR][c] !== "" && values[endR][c] != null) {
                                    endR++;
                                }
                                if (endR > startR) {
                                    const subRange = range.getCell(startR, c).getBoundingRect(range.getCell(endR, c));
                                    subRange.merge(false);
                                }
                                startR = endR + 1;
                            }
                        }
                    } else { // horizontal
                        for (let r = 0; r < rows; r++) {
                            let startC = 0;
                            while (startC < cols) {
                                let endC = startC;
                                while (endC + 1 < cols && values[r][endC] === values[r][endC + 1] && values[r][endC] !== "" && values[r][endC] != null) {
                                    endC++;
                                }
                                if (endC > startC) {
                                    const subRange = range.getCell(r, startC).getBoundingRect(range.getCell(r, endC));
                                    subRange.merge(false);
                                }
                                startC = endC + 1;
                            }
                        }
                    }
                }
                else if (act.action === 'modify_cell_lines') {
                    // Workaround for partial formatting: modify specific lines within a cell
                    range.load("values");
                    await context.sync();

                    const values = range.values;
                    let hasChanges = false;

                    for (let r = 0; r < values.length; r++) {
                        for (let c = 0; c < values[r].length; c++) {
                            const cellValue = values[r][c];
                            if (typeof cellValue === 'string' && cellValue.includes('\n')) {
                                let lines = cellValue.split('\n');
                                if (act.line_index !== undefined && act.line_index >= 0 && act.line_index < lines.length) {
                                    lines[act.line_index] = (act.prefix || "") + lines[act.line_index] + (act.suffix || "");
                                    values[r][c] = lines.join('\n');
                                    hasChanges = true;
                                }
                            }
                        }
                    }
                    if (hasChanges) {
                        range.values = values;
                    }
                }
                else if (act.action === 'set_number_format' && act.format !== undefined) {
                    range.load("rowCount, columnCount");
                    await context.sync();
                    const formatArray = Array(range.rowCount).fill().map(() => Array(range.columnCount).fill(act.format));
                    range.numberFormat = formatArray;
                }
            } catch (err) {
                console.error("Error executing action:", act, err);
            }
        }

        await context.sync();
    });
}

// ============== GEMINI API LOGIC ==============

async function callGeminiAPI(systemInstruction, userPrompt, signal) {
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
        body: JSON.stringify(body),
        signal: signal
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || "Failed to fetch from Gemini API");
    }

    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}
