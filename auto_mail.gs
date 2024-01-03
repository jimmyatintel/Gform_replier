// Replace 'YOUR_WEBHOOK_URL' with the actual URL where you want to send the POST request
const WEBHOOK_URL = 'http://122.116.25.41:5000/webhook';

function onSubmitTrigger(e) {
  const response = e.response;
  var respondentEmail = response.getRespondentEmail();
  const itemResponses = response.getItemResponses();
  
  // Collect form responses or any specific data you want to include in the POST request body
  let formData = {};
  itemResponses.forEach(function(itemResponse) {
    const question = itemResponse.getItem().getTitle();
    const answer = itemResponse.getResponse();
    formData[question] = answer;
  });

  // Make a POST request with the collected form data
  sendPostRequest(formData, respondentEmail);
}

function sendPostRequest(formData, respondentEmail) {
  const payload = {
    email: respondentEmail,
    data: formData
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  // Send the POST request
  UrlFetchApp.fetch(WEBHOOK_URL, options);
}
