/**
 * Clarification Modal Module
 * Handles user clarification prompts during agent execution
 * Supports both simple text questions and structured forms
 */

/**
 * Show clarification modal and return promise that resolves with user's answer
 * @param {string|object} question - Either a text question or structured form object
 */
export function showClarificationModal(question) {
    return new Promise((resolve) => {
        const modal = document.getElementById('clarificationModal');
        const questionElement = document.getElementById('clarificationQuestion');
        const inputElement = document.getElementById('clarificationInput');
        const submitButton = document.getElementById('clarificationSubmit');
        const cancelButton = document.getElementById('clarificationCancel');

        // Detect if question is structured (object with questions array) or simple text
        let isStructured = false;
        let structuredData = null;
        
        try {
            // Try to parse if it's a JSON string
            if (typeof question === 'string' && question.trim().startsWith('{')) {
                structuredData = JSON.parse(question);
                isStructured = structuredData.questions && Array.isArray(structuredData.questions);
            } else if (typeof question === 'object' && question.questions) {
                structuredData = question;
                isStructured = true;
            }
        } catch (e) {
            // Not JSON, treat as simple text
            isStructured = false;
        }

        if (isStructured) {
            // Render structured form
            renderStructuredForm(questionElement, inputElement, structuredData);
        } else {
            // Render simple textarea
            renderSimpleInput(questionElement, inputElement, question);
        }

        // Show modal
        modal.classList.remove('hidden');
        
        // Focus first input
        const firstInput = modal.querySelector('input, textarea, select');
        if (firstInput) firstInput.focus();

        // Submit handler
        const handleSubmit = () => {
            let answer;
            
            if (isStructured) {
                answer = collectStructuredAnswers(structuredData);
                if (!answer) {
                    // Validation failed
                    return;
                }
            } else {
                answer = inputElement.value.trim();
                if (!answer) {
                    inputElement.focus();
                    return;
                }
            }
            
            modal.classList.add('hidden');
            cleanup();
            resolve(answer);
        };

        // Cancel handler
        const handleCancel = () => {
            modal.classList.add('hidden');
            cleanup();
            resolve(null);
        };

        // Enter key handler (only for simple text input)
        const handleKeyDown = (e) => {
            if (e.key === 'Enter' && !e.shiftKey && !isStructured) {
                e.preventDefault();
                handleSubmit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                handleCancel();
            }
        };

        // Cleanup function
        const cleanup = () => {
            submitButton.removeEventListener('click', handleSubmit);
            cancelButton.removeEventListener('click', handleCancel);
            if (!isStructured) {
                inputElement.removeEventListener('keydown', handleKeyDown);
            }
        };

        // Attach event listeners
        submitButton.addEventListener('click', handleSubmit);
        cancelButton.addEventListener('click', handleCancel);
        if (!isStructured) {
            inputElement.addEventListener('keydown', handleKeyDown);
        }
    });
}

/**
 * Render simple text input (original behavior)
 */
function renderSimpleInput(questionElement, inputElement, question) {
    questionElement.innerHTML = `${question}<br><br><small style="color: #6b7280; font-style: italic;">💡 Tip: Select cells in Excel first, then type your answer here. The selection will be captured when you submit.</small>`;
    inputElement.value = '';
    inputElement.style.display = 'block';
}

/**
 * Render structured form with multiple fields
 */
function renderStructuredForm(questionElement, inputElement, structuredData) {
    // Hide the default textarea
    inputElement.style.display = 'none';
    
    // Build form HTML
    let formHTML = `<div class="structured-form">`;
    
    if (structuredData.title) {
        formHTML += `<h4 class="form-title">${structuredData.title}</h4>`;
    }
    
    structuredData.questions.forEach((q, index) => {
        formHTML += `<div class="form-field" data-field-id="${q.id}">`;
        formHTML += `<label class="form-label">${q.label}</label>`;
        
        if (q.hint) {
            formHTML += `<p class="form-hint">${q.hint}</p>`;
        }
        
        switch (q.type) {
            case 'radio':
                q.options.forEach((option, optIndex) => {
                    const isDefault = option === q.default;
                    formHTML += `
                        <div class="radio-option">
                            <input type="radio" 
                                   id="${q.id}_${optIndex}" 
                                   name="${q.id}" 
                                   value="${option}"
                                   ${isDefault ? 'checked' : ''}>
                            <label for="${q.id}_${optIndex}">${option}</label>
                        </div>
                    `;
                });
                break;
                
            case 'select':
                formHTML += `<select id="${q.id}" name="${q.id}" class="form-select">`;
                q.options.forEach(option => {
                    const isDefault = option === q.default;
                    formHTML += `<option value="${option}" ${isDefault ? 'selected' : ''}>${option}</option>`;
                });
                formHTML += `</select>`;
                break;
                
            case 'number':
                formHTML += `<input type="number" 
                                    id="${q.id}" 
                                    name="${q.id}" 
                                    class="form-input"
                                    value="${q.default || ''}"
                                    step="0.1">`;
                break;
                
            case 'text':
            default:
                formHTML += `<input type="text" 
                                    id="${q.id}" 
                                    name="${q.id}" 
                                    class="form-input"
                                    value="${q.default || ''}"
                                    placeholder="${q.hint || ''}">`;
                break;
        }
        
        formHTML += `</div>`;
    });
    
    formHTML += `</div>`;
    
    questionElement.innerHTML = formHTML;
}

/**
 * Collect answers from structured form and format nicely
 */
function collectStructuredAnswers(structuredData) {
    const answers = {};
    const formattedLines = [];
    
    for (const q of structuredData.questions) {
        let value;
        
        if (q.type === 'radio') {
            const selected = document.querySelector(`input[name="${q.id}"]:checked`);
            value = selected ? selected.value : null;
        } else {
            const input = document.getElementById(q.id);
            value = input ? input.value : null;
        }
        
        // Validation: require non-empty values
        if (!value && value !== 0) {
            alert(`Please provide an answer for: ${q.label}`);
            return null;
        }
        
        answers[q.id] = value;
        
        // Format for chat display
        formattedLines.push(`✓ ${q.label}: ${value}`);
    }
    
    // Return formatted string for chat display
    // The backend will receive this as the clarification_answer
    return formattedLines.join('\n');
}

