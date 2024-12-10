Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initializeSpellChecker();
    } else {
        showNotification("Error", "This add-in is designed for Word.", true);
    }
});

// Initialize Spell Checker
function initializeSpellChecker() {
    const runSpellCheckButton = document.getElementById("runSpellCheck");
    runSpellCheckButton.addEventListener("click", runSpellCheck);
}0

// Show or hide progress bar
function toggleProgressBar(show) {
    const progressBar = document.getElementById("progress-bar");
    progressBar.style.display = show ? "block" : "none";
}

// Main function to run spell check
async function runSpellCheck() {
    toggleProgressBar(true); // Show progress bar

    try {
        const documentText = await getDocumentText();
        if (!documentText) {
            showNotification("Info", "No text found in the document.", true);
            toggleProgressBar(false);
            return;
        }

        // Fetch spelling suggestions from the backend once
        const suggestions = await checkSpelling(documentText);

        // Filter suggestions with multiple words
        const filteredSuggestions = filterMultiWordSuggestions(suggestions);

        // Separate misspelled and valid words
        const misspelledWords = Object.keys(filteredSuggestions).filter(
            (word) => !filteredSuggestions[word].includes(word)
        );
        const validWords = getValidWordsFromSuggestions(filteredSuggestions);

        // Highlight misspelled words and clear valid ones
        await highlightMisspelledWords(misspelledWords);
        await clearHighlights(validWords);

        // Update the task pane with misspelled words
        updateTaskpane(filteredSuggestions, misspelledWords);
    } catch (error) {
        showNotification("Error", "Error while running spell checker: " + error.message, true);
    } finally {
        toggleProgressBar(false); // Hide progress bar
    }
}

// Get document text, clean it, and replace punctuation with spaces
async function getDocumentText() {
    return Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        const cleanedText = body.text.replace(/[\p{P}]+/gu, " ").replace(/\d+/g, "").trim();
        return cleanedText;
    }).catch(() => {
        showNotification("Error", "Error retrieving document text.", true);
        return "";
    });
}

// Filter out multi-word suggestions
function filterMultiWordSuggestions(suggestions) {
    const filteredSuggestions = {};
    for (const [word, suggestionList] of Object.entries(suggestions)) {
        if (!/\s/.test(word)) {
            filteredSuggestions[word] = suggestionList;
        }
    }
    return filteredSuggestions;
}

// Ensure backend is called only once per operation
async function checkSpelling(text) {
    try {
        const response = await fetch("http://127.0.0.1:120/suggest", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ input_text: text, top_n: 5 }),
        });

        if (!response.ok) {
            throw new Error("Failed to fetch suggestions from the backend.");
        }

        const data = await response.json();

        // Filter out invalid suggestions (numbers, punctuations)
        const filteredSuggestions = {};
        for (const [word, suggestionList] of Object.entries(data.suggestions || {})) {
            if (!/\d/.test(word)) { // Ignore numbers
                filteredSuggestions[word] = suggestionList.filter(
                    (suggestion) => !/\d/.test(suggestion)
                );
            }
        }

        return filteredSuggestions;
    } catch (error) {
        showNotification("Error", "Error communicating with the backend: " + error.message, true);
        return {};
    }
}

// Extract valid words from suggestions
function getValidWordsFromSuggestions(suggestions) {
    return Object.entries(suggestions)
        .filter(([word, suggestionList]) => suggestionList.length > 0 && suggestionList[0] === word)
        .map(([word]) => word);
}

// Highlight misspelled words in the document
function highlightMisspelledWords(misspelledWords) {
    return Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs.load("items");
        await context.sync();

        for (const paragraph of paragraphs.items) {
            for (const word of misspelledWords) {
                const searchResults = paragraph.search(word, { matchWholeWord: true });
                searchResults.load("items");
                await context.sync();

                searchResults.items.forEach((result) => {
                    result.font.underline = Word.UnderlineType.wave; // Highlight as wavy underline
                });
            }
        }

        await context.sync();
    }).catch(() => {
        showNotification("Error", "Error highlighting misspelled words.", true);
    });
}

// Clear highlights for valid words
function clearHighlights(validWords) {
    return Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs.load("items");
        await context.sync();

        for (const paragraph of paragraphs.items) {
            for (const word of validWords) {
                const searchResults = paragraph.search(word, { matchWholeWord: true });
                searchResults.load("items");
                await context.sync();

                searchResults.items.forEach((result) => {
                    result.font.underline = Word.UnderlineType.none; // Remove highlight
                });
            }
        }

        await context.sync();
    }).catch(() => {
        showNotification("Error", "Error clearing highlights.", true);
    });
}

function updateTaskpane(suggestions, misspelledWords) {
    const taskpaneDiv = document.getElementById("taskpane-content");
    taskpaneDiv.innerHTML = ""; // Clear existing content

    for (const word of misspelledWords) {
        const suggestionList = suggestions[word] || [];
        const container = document.createElement("div");
        container.className = "word-container";

        const label = document.createElement("span");
        label.textContent = word;
        container.appendChild(label);

        const dropdown = document.createElement("select");

        // Add the first placeholder option
        const placeholderOption = document.createElement("option");
        placeholderOption.value = "";
        placeholderOption.textContent = "Click here for suggestions";
        placeholderOption.selected = true; // Pre-select the placeholder
        placeholderOption.disabled = true; // Make the placeholder non-selectable
        dropdown.appendChild(placeholderOption);

        // Add suggestions to the dropdown
        suggestionList.forEach((suggestion) => {
            const option = document.createElement("option");
            option.value = suggestion;
            option.textContent = suggestion;
            dropdown.appendChild(option);
        });

        // Add "Add to Dictionary" option
        const addOption = document.createElement("option");
        addOption.value = "add_to_dictionary";
        addOption.textContent = "Add to Dictionary";
        dropdown.appendChild(addOption);

        // Handle replacement logic
        dropdown.addEventListener("change", async (event) => {
            const newWord = event.target.value;
            if (newWord === "add_to_dictionary") {
                await addToDictionaryHandler(word);
            } else if (newWord) {
                await replaceWordInDocument(word, newWord);
            }
        });

        container.appendChild(dropdown);
        taskpaneDiv.appendChild(container);
    }
}

// Add a word to the dictionary
async function addToDictionaryHandler(word) {
    clearHighlights(word);
    try {
        const response = await fetch("http://127.0.0.1:120/add_word", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ new_word: word }),
        });

        if (!response.ok) {
            throw new Error("Failed to add word to dictionary.");
        }
        showNotification("Success", `"${word}" added to the dictionary.`);
    } catch (error) {
        showNotification("Error", "Error adding word to dictionary: " + error.message, true);
    }
}

// Show notification in the task pane
function showNotification(type, message, isError = false) {
    const notificationArea = document.getElementById("notification-area");
    notificationArea.textContent = `${type}: ${message}`;
    notificationArea.style.color = isError ? "red" : "green";
    notificationArea.style.display = "block";

    setTimeout(() => {
        notificationArea.style.display = "none";
    }, 5000); // Auto-hide after 5 seconds
}

function normalizeText(text) {
    return text.normalize("NFC"); // Use NFC or NFD as needed
}


async function replaceWordInDocument(oldWord, newWord) {
    return Word.run(async (context) => {
        const body = context.document.body;

        const normalizedOldWord = normalizeText(oldWord.trim());
        const normalizedNewWord = normalizeText(newWord.trim());
        
        console.log(`Normalized Old Word: ${normalizedOldWord}`);
        console.log(`Normalized New Word: ${normalizedNewWord}`);

        const searchResults = body.search(normalizedOldWord, {
            matchCase: false,
            matchWholeWord: true,
        });

        searchResults.load("items");
        await context.sync();
        console.log(searchResults.items);
        

        if (searchResults.items.length === 0) {
            showNotification("Info", `"${normalizedOldWord}" not found in the document.`, true);
            return;
        }

        searchResults.items.forEach((result) => {
            clearHighlights(result);
            result.insertText(normalizedNewWord, Word.InsertLocation.replace);
        });

        await context.sync();
        showNotification("Success", `"${oldWord}" replaced with "${newWord}".`);
    }).catch((error) => {
        showNotification("Error", `Error replacing word: ${error.message}`, true);
        console.log(error);
    });
}