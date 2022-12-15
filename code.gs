/* SpreadsheetApp. */
var ssAllImagesURLs = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1e8Q7Yh84XlOrl-w9wVcbOjhoARXUCbaQYj6orMDxSsA/edit#gid=0').getSheetByName('allImagesURLs');
var ssGenerator = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1e8Q7Yh84XlOrl-w9wVcbOjhoARXUCbaQYj6orMDxSsA/edit#gid=484787756').getSheetByName('Generator');
var ssSelectedImagesURLs = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1e8Q7Yh84XlOrl-w9wVcbOjhoARXUCbaQYj6orMDxSsA/edit#gid=484787756').getSheetByName('SelectedImagesURLs');

/* SlidesApp. */
var sa_new = SlidesApp.create("🔮BS Idiom " + new Date());
// Templates.
var saStartSlideTemplate = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1jGixtNa6MOIwv3g6DTycQUKsauX-b_njBsmECZkAHts/edit#slide=id.SLIDES_API1813896228_0');
var saBodySlideTemplate = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1yctjC5IR4pLh9KQ8_8KY2cvfAJGGnTY6sLFzG7q43gc/edit');
var saWordsListSlideTemplate = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1rIblGe-pdgh2MvSuiq4ngGXBOB08owYIil1QlzjOROA/edit');
var saTitleSlideTemplate = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1hyxIV5lzCFAPWIfV6KiWTCYkfKNKA7FWjOCV-vwm8LA/edit');
var ssEndSlideTemplate = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1cJ8-AjPedoPB5VXOdBEQMWEu4q9TZ9L91pdkEwW9PVc/edit'); 

function generatePresentation() {

  var lastRow = lastRowInColumnLetterInSsGenarator('A');
  generateNumBodySlidesFromTemplate(lastRow - 1);

  Logger.log("Last row: " + lastRow)
  // If idiom exist.
  if (lastRow > 1) {
    replaceTextInSlides();
    replaceImageInSlides();
  }
  addWordsListTemplate();
  addTitleSlideTemplate();
  addStartSlideTemplate();
  addEndSlideTemplate()
  
}

/*
function onEdit() {
  var imgURLs = selectValidImagesURLsInSheet();
  writeImagesURLsInSheet(imgURLs);
}
*/
function selectValidImagesURLsInSheet() {

  // Get images url and seve it in imgRange.
  var xLastRow = ssAllImagesURLs.getLastRow();
  var yLastCol = ssAllImagesURLs.getMaxColumns();
  var imgRange = ssAllImagesURLs.getRange(1, 1, xLastRow, yLastCol).getValues();

  // Check images url from ssAllImagesURLs, and then write it in 2D array(selectedImagesURLs).
  var selectedImagesURLs = [];
  for (let y = 0; y < imgRange.length; ++y) {
    let currentPartImages = new Array();
    for (let i = 0; i < imgRange[y].length; ++i) {
      var rx = '^(http|https)://.*OIP.*$';
      if (imgRange[y][i].match(rx)) {
        currentPartImages.push(imgRange[y][i]);
      }
    }
    selectedImagesURLs.push(currentPartImages);
  }
  return selectedImagesURLs
}


function writeImagesURLsInSheet(selectedImagesURLs) {

  // Clear old data.
  var xLastRow = ssSelectedImagesURLs.getLastRow();
  var yLastCol = ssSelectedImagesURLs.getMaxColumns();
  ssSelectedImagesURLs.getRange(2, 2, yLastCol + 1, xLastRow + 1).clearContent();

  // Write Images URLs in ssSalectedImageURLs.
  for (let y = 0; y < selectedImagesURLs.length; ++y) {
    for (let x = 0; x < selectedImagesURLs[y].length; ++x) {
      ssSelectedImagesURLs.getRange(1 + y, 2 + x).setValue(selectedImagesURLs[y][x]);
    }
  }
}

function replaceTextInSlides() {

  //var lastRow = lastRowInColumnLetter('A');
  var lastRow = lastRowInColumnLetterInSsGenarator('A');
  var lastCol = ssGenerator.getLastColumn();
  var new_string = ssGenerator.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var slides = sa_new.getSlides();
  for (var sl = 0; sl < slides.length; ++sl) {
    var shapes = slides[sl].getShapes();
    for (var sh = 0; sh < shapes.length; ++sh) {
      Logger.log(new_string[sl]);

      // Replace texts.
      shapes.forEach(function (shape) {
        shape.getText().replaceAllText('{{idiom}}', new_string[sl][0]);
        shape.getText().replaceAllText('{{meaning}}', new_string[sl][3]);
        shape.getText().replaceAllText('{{example1}}', new_string[sl][4]);
        shape.getText().replaceAllText('{{example2}}', new_string[sl][5]);
        shape.getText().replaceAllText('{{example3}}', new_string[sl][6]);
      });
    }
  }
}

function replaceImageInSlides() {

  var xLastRow = ssSelectedImagesURLs.getLastRow();
  var yLastCol = ssSelectedImagesURLs.getMaxColumns();

  // Get images urls 2D array.
  selectedImages = ssSelectedImagesURLs.getRange(2, 2, yLastCol, xLastRow).getValues();

  var slides = sa_new.getSlides();
  for (var sl = 0; sl < slides.length; ++sl) {
    // Nedd to skip first image.
    let skipLogo = true;
    // Count pasted image.
    let pastedImgInd = 0;
    var elements = slides[sl].getPageElements();

    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].getPageElementType() == "IMAGE") {
        if (skipLogo) {
          skipLogo = false;
        } else {
          try {
            // If Microsoft not bloc urls.
            elements[i].asImage().replace(selectedImages[sl][pastedImgInd]);
            ++pastedImgInd;
          } catch (error) {
            // Else try another image.
            console.error(error);
            ++pastedImgInd;
          }
        }
      }
    }
  }
}

function generateNumBodySlidesFromTemplate(num) {
  Logger.log("Add " + num + " body slides.");
  
  var bodyTemplateSlides = saBodySlideTemplate.getSlides();
  for (let i = 0; i < num; ++i) {
    let r = getRandomINTfromZeroToMAX(bodyTemplateSlides.length);
     Logger.log("Add body slide type #" + r +".");
    sa_new.appendSlide(bodyTemplateSlides[r]);
  }
  // Remove slide which created by default.
  sa_new.getSlides()[0].remove();
}

function addStartSlideTemplate() {
  let startTemplateSlides = saStartSlideTemplate.getSlides();
  let r = getRandomINTfromZeroToMAX(startTemplateSlides.length);
  Logger.log("Add start slide type #" + r +".")
  sa_new.insertSlide(0, startTemplateSlides[r]);
}

function addTitleSlideTemplate(){

  let titleTemplateSlides = saTitleSlideTemplate.getSlides();
  let r = getRandomINTfromZeroToMAX(titleTemplateSlides.length);
  Logger.log("Add title slide type #" + r +".");
  sa_new.insertSlide(0, titleTemplateSlides[r]);
}

function addWordsListTemplate() {
  var wordListTemplateSlides = saWordsListSlideTemplate.getSlides();
  let r = getRandomINTfromZeroToMAX(wordListTemplateSlides.length);
  var lastRow = lastRowInColumnLetterInSsGenarator('A');
  var lastCol = ssGenerator.getLastColumn();
  var new_string = ssGenerator.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var string = "";

  for (let i = 0; i < new_string.length; ++i) {
    string += new_string[i][0] + '\n';
  }
  Logger.log("Add word list slide type #" + r +".");
  sa_new.insertSlide(sa_new.getSlides().length, wordListTemplateSlides[r]);

  var slides = sa_new.getSlides();
  Logger.log(slides.length);

  var shapes = slides[slides.length - 1].getShapes();
  Logger.log(shapes);
  Logger.log(string);

  // Replace texts.
  shapes.forEach(function (shape) {
    shape.getText().replaceAllText('{{words_list}}', string);
  });
}

function addEndSlideTemplate(){
  let endTemlateSlides = ssEndSlideTemplate.getSlides();
  let r = getRandomINTfromZeroToMAX(endTemlateSlides.length);
  sa_new.insertSlide(sa_new.getSlides().length, endTemlateSlides[r]);
}



function lastRowInColumnLetterInSsGenarator(column) {
  var lastRow = ssGenerator.getLastRow() - 1; // values[] array index
  var values = ssGenerator.getRange(column + "1:" + column + (lastRow + 1)).getValues();
  while (lastRow > -1 && values[lastRow] == "") {
    lastRow--;
  }
  if (lastRow == -1) {
    return 0;
  } else {
    return lastRow + 1;
  }
}

function getRandomINTfromZeroToMAX(max){
   return Math.floor(Math.random() * (max));;
}
