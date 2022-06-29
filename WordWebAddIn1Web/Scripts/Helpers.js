﻿/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
    "use strict";
    class SelectedPattern {
        constructor(find, replace) {
            this.find = find;
            this.replace = replace;
        }
    }
    // Create a namespace to hold application-wide settings with primitive data types.
    var WordProcessorApp = window.WordProcessorApp || {  };

    WordProcessorApp.checkedItems = ['longSentence'];


    WordProcessorApp.CurrentSelectedPattern ;
    

    // Ensures that all table data fields are selected in the UI. 
    WordProcessorApp.reinitializeUI = function () {
        
        $.each(WordProcessorApp.CheckBoxElements, function () {
            this.check();  // Use the check method on the Office UI Fabric Checkbox component to check the checkbox.
        });
    }      

    //Serach for a pattern and replace with other Pattern.
    WordProcessorApp.SetResetCurrentPattern = function (event) {
        console.log("inside set pattern");
        var secondaryElement = $(this).parent('.ms-ListItem').children('.ms-ListItem-secondaryText');
        var tertiaryElement = $(this).parent('.ms-ListItem').children('.ms-ListItem-tertiaryText');
        var isChecked = !$(this).parent('.ms-ListItem').hasClass('is-selected');
        if (isChecked) {
            WordProcessorApp.CurrentSelectedPattern =
                new SelectedPattern(secondaryElement[0].textContent, tertiaryElement[0].textContent);                           
        }
        else {
            WordProcessorApp.CurrentSelectedPattern = new SelectedPattern("", "");
        }
        
        console.log(WordProcessorApp.CurrentSelectedPattern.find);          
    }
    WordProcessorApp.SearchandReplace = async function (event) {
        await Word.run(async (context) => {
            console.log("Inside Search and replace");
           
            var findPattern = event.data.find;
            var replacePattern = event.data.replace;
            var type = event.data.type;
            var unicode = event.data.unicode;

            var body = context.document.body;
            var options = Word.SearchOptions.newObject(context);
            options.matchCase = false;
            if (type === "simple") {
                
                var searchResults = context.document.body.search(findPattern, options);
                searchResults.load("length");

                await context.sync();
                //context.load(searchResults, "text");

                for (var i = 0; i < searchResults.items.length; i++) {

                    searchResults.items[i].load('text');
                    await context.sync();
                    var str = searchResults.items[i].text;
                    console.log(str);
                    str = str.replace(findPattern, replacePattern);
                    console.log(str);
                    searchResults.items[i].insertText(str, "Replace");
                }
            }
            else {
                options.matchWildcards = true;
                var searchResults = context.document.body.search(findPattern, options);
                searchResults.load("length");

                await context.sync();
                var regex;
                for (var i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].load('text');
                    await context.sync();
                    var str = searchResults.items[i].text;

                    // / {2,}/;
                    console.log("previous: " + str);
                    if (unicode === "") {
                        regex = new RegExp(findPattern);
                        str = str.replace(regex, replacePattern);
                    }
                    else {
                        regex = new RegExp(replacePattern);
                        str = str.replace(regex, "$1" + unicode);
                    }
                        
                    console.log("repalced: "+str);
                    searchResults.items[i].insertText(str, "Replace");
                }
            }
            

            await context.sync();
        });
        
    }
    // Navigates to a different DIV when users click on the tab bar.
    WordProcessorApp.newPage = function () {
        if (this.id === 'HighlightTab') {
            $('.HighlightSettings').show().siblings().hide();
        }
        else if (this.id === 'FindReplaceTab') {
            $('.FindReplaceSettings').show().siblings().hide();
        }
        else {
            $('.WordOveruseSettings').show().siblings().hide();
        }
    };

        // Gets sales data, and displays the data in a table and chart.
    WordProcessorApp.populateRegexList = function () {        

        $.ajax({
            url: 'FindReplaceRegex.json',
            dataType: "json"
        })
            // Process returned data.
        .then(function (data) {

            var ul = $('#ms-List-Regex');
            data.Patterns.forEach(function (pattern) {

                var iconElement = document.createElement('i');
                iconElement.classList.add("ms-font-l", "ms-fontWeight-light", "ms-fontColor-themePrimary",
                    "ms-Icon", "ms-Icon--Play");
                var divElement = document.createElement('div');
                divElement.className = 'ms-ListItem-action';
                divElement.appendChild(iconElement);

                var divElement3 = document.createElement('div');
                divElement3.className = 'ms-ListItem-actions';

                divElement3.appendChild(divElement);

                var divElement2 = document.createElement('div');
                divElement2.classList.add("ms-ListItem-selectionTarget", "js-toggleSelection");

                //hook checkbox click

                var spanPrimary = document.createElement('span');
                spanPrimary.className = 'ms-ListItem-secondaryText';
                spanPrimary.textContent = pattern.Description;

                //var spanSecondary = document.createElement('span');
                //spanSecondary.className = 'ms-ListItem-secondaryText';
                //spanSecondary.textContent = pattern.Find;

                //var spanTertiary = document.createElement('span');
                //spanTertiary.className = 'ms-ListItem-tertiaryText';
                //spanTertiary.textContent = pattern.Replace;

                var li = document.createElement('li');
                li.classList.add("ms-ListItem", "is-selectable");

                li.appendChild(spanPrimary);
                //li.appendChild(spanSecondary);
                //li.appendChild(spanTertiary);
                li.appendChild(divElement2);
                li.appendChild(divElement3);
               
                //$(divElement2).on('click', function (event) {
                //    $(this).parent('.ms-ListItem').toggleClass('is-selected');
                //});
                ul[0].appendChild(li);
                $(iconElement).on ('click',
                    {
                        find: pattern.Find,
                        replace: pattern.Replace,
                        type: pattern.type,
                        unicode:pattern.unicode
                       
                    }, WordProcessorApp.SearchandReplace);

            });
        });
    }

    WordProcessorApp.setHighlightFilters = function (event) {

        // Get ID of column checkbox that was changed.
        var columnName = event.target.id;
        if (WordProcessorApp.checkedItems.length > 0) {
            if (WordProcessorApp.checkedItems.indexOf(columnName) == -1) {
                WordProcessorApp.checkedItems.push(columnName);
            }
            else {
                WordProcessorApp.checkedItems.pop(columnName);
            }
        } else {
            WordProcessorApp.checkedItems.push(columnName);
        }
        
        
    }
   
    WordProcessorApp.errorHandler = function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    window.WordProcessorApp = WordProcessorApp;
})();