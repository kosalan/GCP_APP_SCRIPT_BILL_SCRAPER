/**
The get_newmails function returns all new emails came within the last 24 hours. 
**/
function get_newmails()
{
  
  var _filterdays = 'newer_than:1d ';
  var _messages = [];
  
  var _conffile = SpreadsheetApp.openByUrl('https://url');
 
  var _senderlist = _conffile.getActiveSheet().getRange("Sheet1!B2:B20").getValues();
  var _subjectlist = _conffile.getActiveSheet().getRange("Sheet1!C2:C20").getValues();
  
  for (var i = 0; i < _senderlist.length; i++)
  {
    var _sender = _senderlist[i][0];
    var _subject = _subjectlist[i][0];
    
    if (_sender != '' && _subject != '')
    {
      var _filtersender = 'from:'+ _sender;
      var _filtersubject = 'subject:' + _subject;      
      var _filterstring = _filterdays + ' ' + _filtersender + ' ' + _filtersubject;
      
      var _threads = GmailApp.search(_filterstring);
    
      for (var j = 0; j < _threads.length; j++)
      {
        for (var k = 0; k < _threads[j].getMessageCount(); k++)
        {
          var _message = _threads[j].getMessages()[k];    
          _messages.push(_message);
        }      
      }      
    }
  }
  return _messages;
}

/**
This function will parse the City of Toronto Hydro Bill.
This will return account number and Amount in a comma seperated format.
**/
function toronto_hydro_parser(_body)
{
  var _today = new Date().toISOString().slice(0,10);
  
  var _indexaccnum = _body.indexOf("Account Number");
  var _beginaccnum = _indexaccnum + 17;
  var _endinaccnum = _beginaccnum + 10;
    
  var _indexamountnum = _body.indexOf("Total Amount Due");
  var _beginamnum = _indexamountnum + 20;
  var _endinamnum = _beginamnum + 8;
   
  var _accountnum = _body.substring(_beginaccnum,_endinaccnum);
  var _amountnum = '$' + _body.substring(_beginamnum,_endinamnum).replace('[', '').replace('*','');
  var _property = get_proptertyname(_accountnum);
  var _service = 'Toronto Electricity';
  
  var _detail = [ _today, _property, _service, _accountnum, _amountnum, 'Not Paid', ' '];
  
  return _detail;
}

/**
This function will parse the Enbridge Gas Bill.
This will return account number and amount in a comma seperated format.
**/
function enbridge_gas_parser(_body)
{
  var _today = new Date().toISOString().slice(0,10);
  
  var _indexaccnum = _body.indexOf("Account Number");
  var _beginaccnum = _indexaccnum + 17;
  var _endinaccnum = _beginaccnum + 16;
    
  var _indexamountnum = _body.indexOf("Current Bill Total");
  var _beginamnum = _indexamountnum + 22;
  var _endinamnum = _beginamnum + 6;
   
  var _accountnum = _body.substring(_beginaccnum,_endinaccnum);
  var _amountnum = '$' + _body.substring(_beginamnum,_endinamnum).replace('[', '').replace('*','');
  var _property = get_proptertyname(_accountnum);
  var _service = 'Enbridge Gas';
  
  var _detail = [ _today, _property, _service, _accountnum, _amountnum, 'Not Paid', ' '];
  
  return _detail;
}

/**
This function will parse the Alectra Utility Bill.
This will return account number and amount in a comma seperated format.
**/
function alectra_electricity_parser(_body)
{
  var _today = new Date().toISOString().slice(0,10);
  
  var _indexamountnum = _body.indexOf("Amount Due");
  var _beginamnum = _indexamountnum + 11;
  var _endinamnum = _beginamnum + 8;
   
  //The Alectra Account Number is hardcoded because email does not include account number.
  var _accountnum = '000000000000';
  var _amountnum = _body.substring(_beginamnum,_endinamnum).replace('[', '').replace('*','');
  var _property = get_proptertyname(_accountnum);
  var _service = 'Alectra Electricity';
  
  var _detail = [ _today, _property, _service, _accountnum, _amountnum, 'Not Paid', ' '];
  
  return _detail;
}

/**
This function will parse the Oshawa PUC Bill
This will return account number and amount in a comma seperated format.
**/
function oshawa_electricity_parser(_body)
{
  var _today = new Date().toISOString().slice(0,10);
  
  var _indexaccnum = _body.indexOf("statement for account");
  var _beginaccnum = _indexaccnum + 22;
  var _endinaccnum = _beginaccnum + 11;
    
  var _indexamountnum = _body.indexOf("Current Charges");
  var _beginamnum = _indexamountnum + 18;
  var _endinamnum = _beginamnum + 6;
   
  var _accountnum = _body.substring(_beginaccnum,_endinaccnum);
  var _amountnum = '$' + _body.substring(_beginamnum,_endinamnum).replace('[', '').replace('*','');
  var _property = get_proptertyname(_accountnum);
  var _service = 'Oshawa Electricity';
  
  var _detail = [ _today, _property, _service, _accountnum, _amountnum, 'Not Paid', ' '];
  
  return _detail;
}

/**
This function returns property name and service provider based on account number
**/
function get_proptertyname(_acctnum)
{
  var _conffile = SpreadsheetApp.openByUrl('https://url');
  var _accountnum = _conffile.getActiveSheet().getRange("Sheet2!C2:C50").getValues();
  
  var _accountnumlist = [];
  
  for (var i = 0; i < _accountnum.length; i++)
  {
    if (_accountnum[i][0] != '')
    {
      _accountnumlist.push(_accountnum[i][0]);
    }
  }
  
  var _properties = _conffile.getActiveSheet().getRange("Sheet2!A2:A50").getValues();
  
  var _proplist = [];
  
  for (var i = 0; i < _properties.length; i++)
  {
    if (_properties.length > 0)
    {
      _proplist.push(_properties[i][0]);
    }
  }
  
  var _acctindex = _accountnumlist.indexOf(_acctnum);
  
  var _property = _proplist[_acctindex];
  
  return _property;
}

/**
Validate email address and get parser
**/
function extract_detail()
{
  var _today = new Date().toISOString().slice(0,10);
  var _detail = '';
  
  var _conffile = SpreadsheetApp.openByUrl('https:url');
  var _logsheet = _conffile.getActiveSheet();
  
  var _messages = get_newmails();
  
  if (_messages.length > 0)
  {
    for (var i = 0; i < _messages.length; i++)
    {
      var _message = _messages[i];
      
      if (_message.getFrom() == 'epost <donotreply-nepasrepondre@notifications.canadapost-postescanada.ca>')
      {
        _detail = [_today, 'unknown', 'unknown', '0000000000', '$00.00', 'Not Paid', 'Check Canada E-Post'];
        _logsheet.appendRow(_detail);
      }
      else if (_message.getFrom() == '<no-reply@markham.ca>')
      {
        _detail = [_today, 'unknown', 'unknown', '0000000000', '$00.00', 'Not Paid', 'Check Markham Property Tax Account'];
        _logsheet.appendRow(_detail);
      }
      else if (_message.getFrom() == 'contactus@torontohydro.com')
      {
        _detail = toronto_hydro_parser(_message.getPlainBody());
        _logsheet.appendRow(_detail);
      }
      else if (_message.getFrom() == 'Enbridge.E-Bill@enbridgegas.com')
      {
        _detail = enbridge_gas_parser(_message.getPlainBody());
        _logsheet.appendRow(_detail);
      }
      else if (_message.getFrom() == 'Alectra Utilities eBilling Service <noreply@alectrautilitiesmail.com>')
      {
        _detail = alectra_electricity_parser(_message.getPlainBody());
        _logsheet.appendRow(_detail);
      }
      else if (_message.getFrom() == '<contactus@opuc.on.ca>')
      {
        _detail = oshawa_electricity_parser(_message.getPlainBody());
        _logsheet.appendRow(_detail);
      }
    }
  } 
}
