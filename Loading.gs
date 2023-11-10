function startProcessing(ss, message) {
  var progressBar = HtmlService.createTemplateFromFile("ProgressBar");     
  progressBar.status = "start";
  progressBar.message = message;
  ss.show(progressBar.evaluate()
      .setWidth(300) // enter the desired width and height here
      .setHeight(150));
  return progressBar;
}

function finishProcessing(ss, progressBar) {
  progressBar.status = "finish";
  ss.show(progressBar.evaluate()
      .setWidth(300)
      .setHeight(150)); 
}