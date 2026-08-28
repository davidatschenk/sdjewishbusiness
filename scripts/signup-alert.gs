// San Diego Jewish Business Networking - contact form signup alerts
// Lives on the "SD Jewish Business Networking - Contact Form (Responses)" sheet.
// Run setup() once. After that it runs by itself on every new response.

const RECIPIENTS = 'david@schenklawfirm.com,dan@diamondlawcorp.com';

function onNewResponse(e) {
  const v = (e && e.namedValues) || {};
  const get = function (key) {
    return v[key] && v[key][0] ? String(v[key][0]).trim() : '';
  };

  const name = get('Name') || '(no name given)';
  const email = get('Email Address');
  const company = get('Company (if you are with a company)');
  const message = get('How can we help?');

  const body = [
    'Someone just signed up through sdjewishbusiness.com.',
    '',
    'Name:    ' + name,
    'Email:   ' + (email || '(none given)'),
    'Company: ' + (company || '(none given)'),
    '',
    'Message:',
    message || '(none given)',
    '',
    '---',
    'Hit reply to answer them directly.',
    'Every response is also saved in the Google Sheet in the group folder.'
  ].join('\n');

  const options = {
    to: RECIPIENTS,
    subject: 'New signup: ' + name + ' - SD Jewish Business Networking',
    body: body
  };
  if (email) {
    options.replyTo = email;
  }

  MailApp.sendEmail(options);
}

function setup() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onNewResponse') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onNewResponse').forSpreadsheet(ss).onFormSubmit().create();

  MailApp.sendEmail(
    RECIPIENTS,
    'Signup alerts are on - SD Jewish Business Networking',
    'Alerts are now switched on.\n\n' +
      'From now on, both of you get an email the moment someone fills out the contact ' +
      'form at sdjewishbusiness.com. The email includes their name, email address, ' +
      'company, and message, and you can reply to it directly.\n\n' +
      'Nothing for you to do.'
  );
}
