// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };

    function loadSampleData(event) {
        var buttonId = event.source.id;
        console.log('testEventObject() called, buttonID: ' + buttonId);
        event.complete();
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            //body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document via button click.",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    };
    function HighlightSentences() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var document = context.document;

            // This variable will keep the search results for the sentences big than 30 words.
            var sentences = document.sentences;
            for (var i = 0; i < sentences.length; i++) {
                var wordsInSentence = sentences[i].Words;
                if (wordsInSentence.Count >= 30) {
                    sentences[i].font.highlightColor = '#FFFF00'; // Yellow
                }                  
            }
            return context.sync()
                .then(context.sync)
                
        })
            .catch(errorHandler);
    }
})();