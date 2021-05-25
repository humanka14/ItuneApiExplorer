function onOpen() {
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('iTunes Menu')
  .addItem('Get Artist Data','displayArtistData')
  .addToUi();
  
}


function displayArtistData() {
  
  // get active sheet
  var ss = SpreadsheetApp.getActiveSheet();
  
  // get artist name from relevant cell
  var artist = ss.getRange(4, 2).getValue();
  
  // run Function to pass artist to iTunes API
  var tracks = calliTunes(artist);
  
  // return data
  var results = tracks['results'];
  var resultsLength = results.length;
  
  // create empty array to push values into
  var output = [];
  
  
  // loop through returned data and extract relevant parts *********************************
  for (var i=0; i<resultsLength; i++) {
    
    var dataItem = results[i];
    
    // get artist name
    var artistName = dataItem['artistName'];
    
    // get album name
    var albumName = dataItem['collectionName'];
    
    // get track name
    var trackName = dataItem['trackName'];
    
    // get artwork image
    var image = '=image("' + dataItem["artworkUrl60"] + '",4,60,60)';
    
    // get hyperlink to sample track
    var hyperlink = '=hyperlink("' + dataItem["previewUrl"] + '","Listen to preview")';
    
    // push items into empty array
    output.push([artistName, albumName, trackName, image, hyperlink]);
    
    // increase height of rows in main Table to better display info
    ss.setRowHeight(i+8,58);
    
  }
  // loop through returned data and extract relevant parts ********************************* 
  
  
  // run Function to sort new array by album name
  var sortedOutput = output.sort(sortByAlbum);
  var sortedOutputLength = sortedOutput.length;
  
  /*
  add an index number to the array
  https://www.w3schools.com/jsref/jsref_foreach.asp
  */
  sortedOutput.forEach(function(dataItem, i) {
    dataItem.unshift(i+1);
  });

  // get new length of array with added index numbers
  var sortedOutputNewLength = sortedOutput.length;
  Logger.log('sortedOutputNewLength is: ' + sortedOutputNewLength);
  
  // clear previous Sheet Table content
  ss.getRange(8, 1, 500, 6).clearContent();
  
  // paste in the data
  ss.getRange(8, 1, sortedOutputNewLength, 6).setValues(sortedOutput);
  
  // add some Table formatting
  ss.getRange(8,1,500,6).setVerticalAlignment("middle");
  ss.getRange(8,5,500,1).setHorizontalAlignment("center");
  ss.getRange(8,2,sortedOutputNewLength,3).setWrap(true);
  
}


function calliTunes(artist) {
  
  // iTunes API Url
  var url = 'https://itunes.apple.com/search?term=' + artist + '&limit=200';
  
  // fetch API Url
  var response = UrlFetchApp.fetch(url);
  
  // parse the JSON reply so we can then sift through relevant parts
  var json = response.getContentText();
  var data = JSON.parse(json);
 
  // comment out this line if want to see below extraction working
  return data;
}


function sortByAlbum(a, b) {
  
  // in case album name undefined (which would break script)
  var albumA = (a[1]) ? a[1] : 'Not known';
  var albumB = (b[1]) ? b[1] : 'Not known';
  
  if (albumA < albumB) { 
    return -1; 
  }
  else if (albumA > albumB) {
    return 1;
  }
  
  // otherwise names are equal
  return 0;
  
}