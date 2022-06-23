
(function () {
    "use strict";
    const WordLimitInSentence = 30;
    const CommaLimitInSentence = 3;
    const WordLimitShortSentence = 15;
    const WordLimitInBrackets = 20;
    const BracketMatchingPattern = /\{([^}]+)\}/gm;
    
    var WordProcessorApp = window.WordProcessorApp || {};

    var HighlightedSentences = [];
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            //var element = document.querySelector('.MessageBanner');
            //messageBanner = new fabric['MessageBanner'].MessageBanner(element);
            //messageBanner.hideBanner();

            WordProcessorApp.messageBanner = $('.MessageBanner').map(function () {
                return new fabric['MessageBanner'](this);
            });
            WordProcessorApp.messageBanner.hide();

            $('.ms-ListItem').ListItem();
            

            $('#hightlightbutton-text').text("Highlight");
            $('#highlightbutton-desc').text("Click the button to highlights the matched sentences in the document.");

            $('#HighlightTab').click(WordProcessorApp.newPage);
            $('#FindReplaceTab').click(WordProcessorApp.newPage);
            $('#WordOveruseTab').click(WordProcessorApp.newPage);

            $('.column-selector').children('input[type="checkbox"]').change(WordProcessorApp.setHighlightFilters);

            
            //// If not using Word 2016, use fallback logic.
            ////if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
            ////    $("#template-description").text("This sample displays the selected text.");
            ////    $('#highlightbutton-text').text("Display!");
            ////    $('#highlightbutton-desc').text("Display the selected text");
                
            ////    $('#highlight-button').click(displaySelectedText);
            ////    return;
            ////}

            
            WordProcessorApp.populateRegexList();
            $('.ms-ListItem').children('.ms-ListItem-selectionTarget').click(
                {
                    //find: $(this).siblings('.ms-ListItem-primaryText').first().innerText,
                    //replace: $(this).siblings('.ms-ListItem-secondaryText').first().innerText,
                    isSelected: $(this).children('.ms-ListItem-selectionTarget').first().hasClass('js-toggleSelection')
                }, WordProcessorApp.SetResetCurrentPattern);

            WordProcessorApp.populateWordOveruse();

            WordProcessorApp.CheckBoxElements = $(".column-selector").map(function () {
                return new fabric['CheckBox'](this);
            });
            

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(HighlightSentences); 
            $('#clear-button').click(ClearHighlights);

            $('.ms-ListItem-action').click({
                    
                    from: $(this).attr('class')
            }, WordProcessorApp.SearchandReplace);

            
        });
    };
    async function ClearHighlights() {
        await Word.run(async (context) => {

            let paragraphs = context.document.body.paragraphs;
            paragraphs.load("$none"); // No properties needed.

            await context.sync();

            for (var p = 0; p < paragraphs.items.length; p++) {
                let sentences = paragraphs.items[p]
                    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
                sentences.load("text");
                await context.sync();
               
                for (let i = 0; i < sentences.items.length; i++) {
                    var foundSentence = HighlightedSentences.filter(function (item) {
                        return (item.text == sentences.items[i].text);
                    });
                    if (foundSentence.length>0) {
                        sentences.items[i].font.highlightColor = "#FFFFFF";
                    }
                    
                }                
            }
            await context.sync();
            
            });
    }
    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightSelectedWord(selectedWord) {
        Word.run(function (context) {
            var range = context.document.getSelection();
            var searchResults;
            context.load(range, 'text');

            return context.sync()
                .then(function () {
                   
                    // Queue a search command.
                    searchResults = range.search(selectedWord, { matchCase: false, matchWholeWord: true });
                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items.forEach((item) => {
                        item.font.highlightColor = '#FFFF00';
                        item.font.size = 14;
                    });                   
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }
    

    async function getParagraph() {
        await Word.run(async (context) => {
            // The collection of paragraphs of the current selection returns the full paragraphs contained in it.
            let paragraph = context.document.getSelection().paragraphs.getFirst();
            paragraph.load("text");

            await context.sync();
            console.log(paragraph.text);
        });
    }
    
    async function HighlightSentences() {
        await Word.run(async (context) => {

            let found = false;
            // Gets the complete sentence (as range) associated with the insertion point.
            let paragraphs = context.document.body.paragraphs;
            paragraphs.load("$none"); // No properties needed.

            await context.sync();

            for (var p = 0; p < paragraphs.items.length; p++) {
                let sentences = paragraphs.items[p]
                    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
                sentences.load("text");
                await context.sync();
                let previousCount = await GetWordCount(sentences.items[0].text);
                let k = 0;
                for (let i = 0; i < sentences.items.length; i++) {

                    //check if the sentence has more than 30 words
                    if (WordProcessorApp.checkedItems.indexOf("longSentence") !== -1) {
                        var isOverLimit = await IsWordCountInSentenceOverLimit(sentences.items[i].text);
                        if (isOverLimit) {
                            sentences.items[i].font.highlightColor = '#FFFF00'; // Yellow
                            found = true;
                            HighlightedSentences.push(sentences.items[i]);
                        }
                    }

                    //Check if sentences has multiple commas
                    if (WordProcessorApp.checkedItems.indexOf("complicatedSentence") !== -1) {
                        var hasMultiCommas = await HasSentenceMultipleCommas(sentences.items[i].text);
                        if (hasMultiCommas) {
                            sentences.items[i].font.highlightColor = '#FFFFE0';
                            found = true;
                            //add to the collection
                            HighlightedSentences.push(sentences.items[i]);
                        }
                    }
                    //Check if  consecutive sentences are too short.
                    if (WordProcessorApp.checkedItems.indexOf("shortSentence") !== -1 && sentences.items.length > 1) {
                        if (i < sentences.items.length-1)
                            k = i + 1;
                        let wordCount = await GetWordCount(sentences.items[k].text);
                        if (wordCount < WordLimitShortSentence && previousCount < WordLimitShortSentence) {
                            if (k == 1) {
                                sentences.items[0].font.highlightColor = '#DD9F00';
                                HighlightedSentences.push(sentences.items[0]);
                                found = true;//Golden
                                //console.log(sentences.items[0].text);
                            }
                            sentences.items[k - 1].font.highlightColor = '#DD9F00'; //Golden
                            sentences.items[k].font.highlightColor = '#DD9F00';
                            found = true;
                            HighlightedSentences.push(sentences.items[k]);
                            HighlightedSentences.push(sentences.items[k-1]);

                        }
                        previousCount = wordCount;
                    }
                    //check if sentences have words in brackets over limit.
                    if (WordProcessorApp.checkedItems.indexOf("bracketSentences") !== -1) {
                        let found = sentences.items[i].text.match(BracketMatchingPattern);
                        if (found) {
                            var aboveLimit = found.filter(IsAboveWordLimitThreashold);
                            if (aboveLimit.length > 0) {
                                found = true;
                                sentences.items[k].font.highlightColor = '#DD9F30';
                                HighlightedSentences.push(sentences.items[k]);
                            }

                        }
                    }
                }
            }
           
            if (!found) {
                showNotification("Highlight", "No match found, nothing to highlight.");
            }
        });
    }

    function IsAboveWordLimitThreashold(matchedItem) {
        var res = GetWordCount(matchedItem);
        console.log(res);
        return res >= WordLimitInBrackets; 
    }

    async function setup() {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.clear();
            body.insertParagraph(
                "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
                "Start"
            );
            body.paragraphs
                .getLast()
                .insertText(
                    "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.",
                    "Replace"
                );
        });
    }
    async function HasSentenceMultipleCommas(sentence) {
        let text = [];
        let paragraph = sentence.trim();
        if (paragraph) {
            paragraph.split(",").forEach((term) => {
                let currentTerm = term.trim();
                if (currentTerm) {
                    text.push(currentTerm);
                }
            });
        }
        if (text.length > CommaLimitInSentence) {            
            return true;
        }
        else
            return false;

    }

    async function IsWordCountInSentenceOverLimit(sentence) {
        let text = [];        
        let paragraph = sentence.trim();
        if (paragraph) {
            paragraph.split(" ").forEach((term) => {
                let currentTerm = term.trim();
                if (currentTerm) {
                    text.push(currentTerm);
                }
            });
        }
        if (text.length > WordLimitInSentence) {
            return true;
        }
        else
            return false;
        
    }
    WordProcessorApp.populateWordOveruse = async function (event) {
        await getDistinctWordCount();
    }

    function GetWordCount(sentence) {
        let text = [];
        let paragraph = sentence.trim();
        if (paragraph) {
            paragraph.split(" ").forEach((term) => {
                let currentTerm = term.trim();
                if (currentTerm) {
                    text.push(currentTerm);
                }
            });
        }
        console.log(text.length);
        return text.length;

    }
    async function getDistinctWordCount() {
        await Word.run(async (context) => {
            let paragraphs = context.document.body.paragraphs;
            paragraphs.load("text");
            await context.sync();

            let text = [];
            paragraphs.items.forEach((item) => {
                let paragraph = item.text.trim();
                if (paragraph) {
                    paragraph.split(/\s+/).forEach((term) => {
                        let currentTerm = term.trim();
                        if (currentTerm) {
                            text.push(currentTerm);
                        }
                    });
                }
            });

            let makeTextDistinct = new Set(text);
            let distinctText = Array.from(makeTextDistinct);
            let allSearchResults = [];

            for (let i = 0; i < distinctText.length; i++) {
                let results = context.document.body.search(distinctText[i], { matchCase: true, matchWholeWord: true });
                results.load("text");

                // Map search term with its results.
                let correlatedResults = {
                    searchTerm: distinctText[i],
                    hits: results
                };

                allSearchResults.push(correlatedResults);
            }

            await context.sync();

            // Display counts.
            var ul = $('#ms-List-Overuse');
            allSearchResults.sort((a,b) => {                
                    return (b.hits.items.length - a.hits.items.length);
            }).forEach((result) => {
                let length = result.hits.items.length;
                //populate wordoveruse list.
                //var iconElement = document.createElement('i');
                //iconElement.classList.add("ms-font-xxl", "ms-fontWeight-light", "ms-fontColor-themePrimary",
                //    "ms-Icon", "ms-Icon--Play");
                //var divElement = document.createElement('div');
                //divElement.className = 'ms-ListItem-action';
                //divElement.appendChild(iconElement);

                var spanSecondary = document.createElement('span');
                spanSecondary.className = 'ms-ListItem-action';
                spanSecondary.textContent = length;


                var divElement3 = document.createElement('div');
                divElement3.className = 'ms-ListItem-actions';

                

                divElement3.appendChild(spanSecondary);
                //divElement3.appendChild(divElement);
                
                var divElement2 = document.createElement('div');
                divElement2.classList.add("ms-listitem-selectiontarget");//js-toggleSelection

                var spanPrimary = document.createElement('span');
                spanPrimary.className = 'ms-ListItem-primaryText';
                spanPrimary.textContent = result.searchTerm;

               
                var li = document.createElement('li');
                li.classList.add("ms-ListItem", "is-selectable");

                li.appendChild(spanPrimary);
                //li.appendChild(spanSecondary);
                
                li.appendChild(divElement2);
                li.appendChild(divElement3);

                ul[0].appendChild(li);
                //console.log("Search term: " + result.searchTerm + " => Count: " + length);
            });
        });
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        }
    }
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        //WordProcessorApp.messageBanner.
        WordProcessorApp.messageBanner.toggleExpansion();
    }

    window.WordProcessorApp = WordProcessorApp;
   
})();
