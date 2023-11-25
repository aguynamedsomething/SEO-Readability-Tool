function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Readability Analysis')
    .addItem('Show Analysis', 'showAnalysis')
    .addItem('Custom Word Analysis', 'showCustomWordAnalysisDialog')
    .addItem('Remove Highlights', 'removeHighlights')
    .addToUi();
}

function showCustomWordAnalysisDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
    .setWidth(400)
    .setHeight(300);
  DocumentApp.getUi().showModalDialog(html, 'Custom Word Analysis');
}

function showAnalysis() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();

  var totalWordCount = countWords(text);
  var sentenceLength = calculateAverageSentenceLength(text);
  var paragraphLength = calculateAverageParagraphLength(body);
  var readabilityScore = calculateFleschReadingEase(text);

  clearAllHighlights(body);
  highlightLongSentences(body, 20);
  highlightLongHeaders(body, 300);

  var ui = DocumentApp.getUi();
  var message = "Total Word Count: " + totalWordCount + "\n" +
                "Average Sentence Length: " + sentenceLength.toFixed(2) + " words\n" +
                "Average Paragraph Length: " + paragraphLength.toFixed(2) + " words\n" +
                "Flesch Reading Ease Score: " + readabilityScore.toFixed(2) + "\n";

  ui.alert('Readability Analysis', message, ui.ButtonSet.OK);
}


function calculateAverageSentenceLength(text) {
  var sentences = text.match(/[^\.!\?]+[\.!\?]+/g);
  if (!sentences) return 0;
  var totalLength = sentences.reduce(function(acc, sentence) {
    return acc + sentence.split(/\s+/).length;
  }, 0);
  return totalLength / sentences.length;
}

function calculateAverageParagraphLength(body) {
  var paragraphs = body.getParagraphs();
  var totalLength = 0;
  paragraphs.forEach(function(paragraph) {
    totalLength += paragraph.getText().split(/\s+/).length;
  });
  return totalLength / paragraphs.length;
}

function calculateFleschReadingEase(text) {
  var sentences = text.match(/[^\.!\?]+[\.!\?]+/g) || [];
  var wordCount = countWords(text);
  var syllableCount = countSyllables(text);
  var sentenceCount = sentences.length;

  return 206.835 - 1.015 * (wordCount / sentenceCount) - 84.6 * (syllableCount / wordCount);
}

function countWords(text) {
  if (!text) return 0;
  text = text.trim();
  return text === "" ? 0 : text.split(/\s+/).length;
}

function countSyllables(text) {
  text = text.toLowerCase();
  if (text.length === 0) return 0;
  text = text.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '');
  text = text.replace(/^y/, '');
  return text.match(/[aeiouy]{1,2}/g).length;
}

function highlightLongSentences(body, wordLimit) {
  var sentences = body.getText().match(/[^\.!\?]+[\.!\?]+/g) || [];
  sentences.forEach(function(sentence) {
    if (sentence.split(/\s+/).length > wordLimit) {
      var searchResult = body.findText(sentence.trim());
      while (searchResult !== null) {
        var foundText = searchResult.getElement().asText();
        var startOffset = searchResult.getStartOffset();
        var endOffset = searchResult.getEndOffsetInclusive();
        foundText.setBackgroundColor(startOffset, endOffset, '#FFFF00'); // Yellow highlight
        searchResult = body.findText(sentence.trim(), searchResult);
      }
    }
  });
}

function highlightLongHeaders(body, wordLimit) {
  var paragraphs = body.getParagraphs();
  var count = 0;
  paragraphs.forEach(function(paragraph, index) {
    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL && paragraphs.length > index + 1) {
      var text = '';
      for (var i = index + 1; i < paragraphs.length && paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.NORMAL; i++) {
        text += paragraphs[i].getText() + ' ';
      }
      if (countWords(text) > wordLimit) {
        paragraph.setBackgroundColor('#FF9999'); // Light red highlight
      }
    }
  });
}

function removeHighlights() {
  var body = DocumentApp.getActiveDocument().getBody();
  clearAllHighlights(body);
}

function clearAllHighlights(body) {
  var paragraphs = body.getParagraphs();
  paragraphs.forEach(function(paragraph) {
    var text = paragraph.editAsText();
    text.setBackgroundColor(null);
  });
}

function highlightWordInDocument(word, color, exactLength) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var searchResult = body.findText(word);

  while (searchResult !== null) {
    var element = searchResult.getElement().asText();
    var startOffset = searchResult.getStartOffset();
    var endOffset = searchResult.getEndOffsetInclusive();

    // Extract the word and its surrounding characters
    var extendedWord = extractExtendedWord(element.getText(), startOffset, endOffset);

    // Check if the word is a standalone word and matches the exact length
    if (extendedWord === ' ' + word + ' ' && word.length === exactLength) {
      element.setBackgroundColor(startOffset, endOffset, color);
    }

    searchResult = body.findText(word, searchResult);
  }
}

function extractExtendedWord(text, start, end) {
  var before = (start > 0) ? text.charAt(start - 1) : ' ';
  var after = (end < text.length - 1) ? text.charAt(end + 1) : ' ';
  return before + text.substring(start, end + 1) + after;
}

function countWordsInText(text, minLength, minOccurrences) {
  var words = text.toLowerCase().match(/\b\w+\b/g);
  var wordCount = {};

  words.forEach(function(word) {
    if (word.length >= minLength) {
      wordCount[word] = (wordCount[word] || 0) + 1;
    }
  });

  return wordCount;
}

function countWordsInDocument() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var elements = body.getNumChildren();
  var inHeader = false;
  var sectionWordCount = 0;
  var results = [];
  var currentHeader = "Start of Document";

  for (var i = 0; i < elements; i++) {
    var element = body.getChild(i);
    var type = element.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        if (inHeader && sectionWordCount > 0) {
          results.push({ header: currentHeader, wordCount: sectionWordCount });
          sectionWordCount = 0;
        }
        currentHeader = paragraph.getText();
        inHeader = true;
      } else if (inHeader) {
        sectionWordCount += countWords(paragraph.getText());
      }
    }
  }

  if (sectionWordCount > 0) {
    results.push({ header: currentHeader, wordCount: sectionWordCount });
  }

  return results;
}

function customWordAnalysis(minLength, minOccurrences) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();
  var wordCount = countWordsInText(text, minLength, minOccurrences);

  for (var word in wordCount) {
    if (wordCount[word] >= minOccurrences) {
      highlightWordInDocument(word, getRandomColor(), minLength);
    }
  }

  return Object.keys(wordCount).map(function(word) {
    return word + ": " + wordCount[word];
  });
}


function getRandomColor() {
  var letters = '0123456789ABCDEF';
  var color = '#';
  for (var i = 0; i < 6; i++) {
    color += letters[Math.floor(Math.random() * 16)];
  }
  return color;
}
