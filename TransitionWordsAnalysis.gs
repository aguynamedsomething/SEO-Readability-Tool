function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('SEO Readability Tool')
    .addItem('Show Analysis', 'showAnalysis') // Assuming you have a function named 'showAnalysis'
    .addItem('Custom Word Analysis', 'showCustomWordAnalysisDialog')
    .addItem('Remove Highlights', 'removeHighlights')
    .addSeparator()
    .addItem('Set Language for Transition Words', 'setLanguage')
    .addItem('Analyze Transition Words', 'analyzeDocument')
    .addToUi();
}


var currentLanguage; // No default language is set

var transitionWords = {
  "German": ["weil", "doch", "mit anderen worten", "so dass"],
  "Dutch": ["omdat", "maar", "net als", "ter conclusie"],
  "French": ["car", "toutefois", "si bien que", "en raison de"],
  "Spanish": ["porque", "pero", "a causa de", "sin embargo"],
  "Italian": ["perché", "però", "a causa", "in sentesi"],
  "Portuguese": ["pois", "contudo", "por causa de", "em suma"],
  "Russian": ["потому", "однако", "потому что", "в итоге"],
  "Polish": ["ponieważ", "jednak", "z uwagi że", "w podsumowaniu"],
  "Catalan": ["perquè", "resumint", "pel que", "en a resum"],
  "Swedish": ["emellertid", "men", "i syfte att", "för att sammanfatta"],
  "Hungarian": ["mivel", "azonban", "ahhoz hogy", "más szóval"],
  "Arabic": ["بينما", "حيثما", "هكذا", "كذلك", "كما"],
  "Hebrew": ["למרות", "בשביל", "כגון", "מלבד", "מפאת"],
  "Indonesian": ["berikut", "kedua", "terutamanya", "terdahulu", "contohnya"],
  "Turkish": ["fakat", "ama", "çünkü", "yüzünden", "topyekun"],
  'Japanese': ["だから", "そのため", "第一に", "具体的には"],
  'English': ['therefore', 'however', 'moreover', 'furthermore', 'consequently', 
              'likewise', 'subsequently', 'indeed', 'thus', 'meanwhile', 
              'nonetheless', 'alternatively', 'similarly', 'otherwise', 'finally', 
              'in conclusion', 'in summary']
};

function setLanguage() {
  var html = HtmlService.createHtmlOutputFromFile('LanguagePicker')
    .setWidth(300)
    .setHeight(200);
  DocumentApp.getUi().showModalDialog(html, 'Select Language');
}

function updateLanguage(language) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('currentLanguage', language);
  Logger.log('Language set to: ' + language);
}


function analyzeDocument() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentLanguage = scriptProperties.getProperty('currentLanguage');
  
  if (!currentLanguage) {
    DocumentApp.getUi().alert('Please set a language first.');
    return;
  }

  var doc = DocumentApp.getActiveDocument();
  var text = doc.getBody().getText();
  var sentences = text.match(/[^\.!\?]+[\.!\?]+/g) || [];
  var transitionWordCount = 0;

  sentences.forEach(function(sentence) {
    if (containsTransitionWord(sentence, currentLanguage)) {
      transitionWordCount++;
    }
  });

  var percentage = sentences.length > 0 ? (transitionWordCount / sentences.length) * 100 : 0;
  var message = 'Percentage of sentences with transition words: ' + percentage.toFixed(2) + '%.\n';

  if (percentage < 30) {
    message += 'Consider adding more transition words. Less than 30% of your sentences contain transition words.';
  } else {
    message += 'Good job! Your document contains a sufficient number of transition words.';
  }

  DocumentApp.getUi().alert('Transition Words Analysis', message, DocumentApp.getUi().ButtonSet.OK);
}

function containsTransitionWord(sentence, language) {
  if (!transitionWords[language] || !Array.isArray(transitionWords[language])) {
    Logger.log('Transition words for ' + language + ' are not defined or not an array.');
    return false;
  }

  return transitionWords[language].some(function(word) {
    return sentence.toLowerCase().includes(word);
  });
}

