Office.onReady((info) => {
    if (info.host === Office.HostType.Word) initializeSpellChecker();
    else showNotification("Error", "This add-in is designed for Word.", true);
});

function initializeSpellChecker() {
    document.getElementById("runSpellCheck").addEventListener("click", runSpellCheck);
}

function toggleProgressBar(show) {
    document.getElementById("progress-bar").style.display = show ? "block" : "none";
}

async function runSpellCheck() {
    toggleProgressBar(true);
    try {
        const documentText = await getDocumentText();
        if (!documentText) {
            showNotification("Info", "No text found in the document.", true);
            toggleProgressBar(false);
            return;
        }
        const suggestions = await checkSpelling(documentText);
        const filteredSuggestions = filterMultiWordSuggestions(suggestions);
        const misspelledWords = Object.keys(filteredSuggestions).filter(
            (word) => !filteredSuggestions[word].includes(word)
        );
        const validWords = getValidWordsFromSuggestions(filteredSuggestions);
        await highlightMisspelledWords(misspelledWords);
        await clearHighlights(validWords);
        updateTaskpane(filteredSuggestions, misspelledWords);
    } catch (error) {
        showNotification("Error", "Error while running spell checker.", true);
    } finally {
        toggleProgressBar(false);
    }
}

async function getDocumentText() {
    return Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text.replace(/[\p{P}]+/gu, " ").replace(/\d+/g, "").trim();
    }).catch(() => {
        showNotification("Error", "Error retrieving document text.", true);
        return "";
    });
}

function filterMultiWordSuggestions(suggestions) {
    const filteredSuggestions = {};
    for (const [word, suggestionList] of Object.entries(suggestions)) {
        if (!/\s/.test(word)) filteredSuggestions[word] = suggestionList;
    }
    return filteredSuggestions;
}

async function checkSpelling(text) {
    try {
        const response = await fetch("https://beemne.pythonanywhere.com/suggest", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ input_text: text, top_n: 5 }),
        });
        if (!response.ok) throw new Error("Failed to fetch suggestions from the backend.");
        const data = await response.json();
        const filteredSuggestions = {};
        for (const [word, suggestionList] of Object.entries(data.suggestions || {})) {
            if (!/\d/.test(word)) {
                filteredSuggestions[word] = suggestionList.filter((s) => !/\d/.test(s));
            }
        }
        return filteredSuggestions;
    } catch {
        showNotification("Error", "Error communicating with the backend.", true);
        return {};
    }
}

function getValidWordsFromSuggestions(suggestions) {
    return Object.entries(suggestions)
        .filter(([_, suggestionList]) => suggestionList[0] === _)
        .map(([word]) => word);
}

function highlightMisspelledWords(misspelledWords) {
    return Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs.load("items");
        await context.sync();
        for (const paragraph of paragraphs.items) {
            for (const word of misspelledWords) {
                const searchResults = paragraph.search(word, { matchWholeWord: false });
                searchResults.load("items");
                await context.sync();
                searchResults.items.forEach((result) => {
                    result.font.underline = Word.UnderlineType.wave;
                });
            }
        }
        await context.sync();
    });
}

async function clearHighlights(validWords) {
    return Word.run(async (context) => {
        const body = context.document.body;
        for (const word of validWords) {
            const searchResults = body.search(word, { matchWholeWord: false });
            searchResults.load("items");
            await context.sync();
            searchResults.items.forEach((result) => {
                result.font.underline = Word.UnderlineType.none;
            });
        }
        await context.sync();
    });
}

function updateTaskpane(suggestions, misspelledWords) {
    const taskpaneDiv = document.getElementById("taskpane-content");
    taskpaneDiv.innerHTML = "";
    misspelledWords.forEach((word) => {
        const container = document.createElement("div");
        container.className = "word-container";
        const label = document.createElement("span");
        label.textContent = word;
        container.appendChild(label);
        const dropdown = document.createElement("select");
        const placeholderOption = document.createElement("option");
        placeholderOption.value = "";
        placeholderOption.textContent = "Click here for suggestions";
        placeholderOption.selected = true;
        placeholderOption.disabled = true;
        dropdown.appendChild(placeholderOption);
        suggestions[word].forEach((suggestion) => {
            const option = document.createElement("option");
            option.value = suggestion;
            option.textContent = suggestion;
            dropdown.appendChild(option);
        });
        dropdown.addEventListener("change", async (event) => {
            const newWord = event.target.value;
            if (newWord) await replaceWordInDocument(word, newWord);
        });
        container.appendChild(dropdown);
        taskpaneDiv.appendChild(container);
    });
}


function removeWordFromTaskpane(word) {
    const taskpaneDiv = document.getElementById("taskpane-content");
    const wordContainers = Array.from(taskpaneDiv.getElementsByClassName("word-container"));
    wordContainers.forEach((container) => {
        if (container.querySelector("span")?.textContent === word) {
            taskpaneDiv.removeChild(container);
        }
    });
}

function showNotification(type, message, isError = false) {
    const notificationArea = document.getElementById("notification-area");
    notificationArea.textContent = `${type}: ${message}`;
    notificationArea.style.color = isError ? "red" : "green";
    notificationArea.style.display = "block";
    setTimeout(() => (notificationArea.style.display = "none"), 5000);
}

function normalizeText(text) {
    return text.normalize("NFC");
}

async function replaceWordInDocument(oldWord, newWord) {
    await Word.run(async (context) => {
        const body = context.document.body;
        const searchResults = body.search(normalizeText(oldWord.trim()), {
            matchCase: false,
            matchWholeWord: false,
        });
        searchResults.load("items");
        await context.sync();
        searchResults.items.forEach((result) => {
            result.insertText(normalizeText(newWord.trim()), Word.InsertLocation.replace);
            result.font.underline = Word.UnderlineType.none;
        });
        await context.sync();
        removeWordFromTaskpane(oldWord);
        showNotification("Success", `"${oldWord}" replaced with "${newWord}".`);
    });
}
