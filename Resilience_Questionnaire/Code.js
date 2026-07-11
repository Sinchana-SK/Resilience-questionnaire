function doGet() {
  return HtmlService.createTemplateFromFile('Assessment')
    .evaluate()
    .setTitle('Resilience Assessment System');
}

function submitAssessment(payload) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Responses");

  const candidate = payload.candidate;

  const scores = payload.responses.map(r => r.score);

  sheet.appendRow([
    new Date(),
    candidate.fullName,
    candidate.email,
    candidate.regNumber,
    candidate.department,
    candidate.programme,
    ...scores,
    payload.totalScore
  ]);
}

function sendEmailReport(payload) {

  const email = payload.candidate.email;

  const subject = "Resilience Assessment Report";

  const score = payload.totalScore;

  let interpretation = "";

  if (score <= 37) {
    interpretation = `A developing level of resilience. Your score indicates that, although you may not always feel at the mercy of events, you would in fact benefit significantly from developing aspects of your behaviour.`;
  }
  else if (score <= 43) {
    interpretation = `An established level of resilience. Your score indicates that you may occasionally have tough days when you can't quite make things go your way, but you rarely feel ready to give up.`;
  }
  else if (score <= 48) {
    interpretation = `A strong level of resilience. Your above-average score indicates that you are pretty good at rolling with the punches and you have an impressive track record of turning setbacks into opportunities.`;
  }
  else {
    interpretation = `An exceptional level of resilience. Your score indicates that you are very resilient most of the time and rarely fail to bounce back—whatever life throws at you. You believe in making your own luck.`;
  }

  const body =
`Greetings!

Thank you for taking part in the "Resilience Assessment", organized by Centre For Academic and Professional Support.

Your report has been attached herewith. It contains detailed information about your assessment scores and interpretation for the same.

Total Score: ${payload.totalScore}/60

Interpretation: ${interpretation};

If you have any queries regarding the results, feel free to revert to this email ID. We will make all possible attempts to get back to you at the earliest.

Warm regards,
Psychometric Assessment Wing
Centre for Academic and Professional Support
CHRIST (Deemed to be University)
Bengaluru-560029
`;



  MailApp.sendEmail(email, subject, body);
}