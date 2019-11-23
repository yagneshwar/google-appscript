function code()
{

  var threads = GmailApp.getInboxThreads();
  var caseid;
  var casedt;
  //var msgid=[];
  //var k=0;
  var link;
  var agent;
  var appealsspread=SpreadsheetApp.openById('paste sheet id');
  var appealactive=appealsspread.getSheetByName("tab name");
  var appealdecision=appealsspread.getSheetByName("tab name");
  var work;
  var queue;
  var alignment;
  for (var i = 0; i < threads.length; i++) {
      if (threads[i].getFirstMessageSubject()=="paste here email subject line")
         {
           
           var messleng = threads[i].getMessageCount();
           for (var l=0; l<messleng ; l++)
           {
             var t = threads[i].getMessages()[l];
             if(t.isStarred()==false)
             {
             var casestr= t.getBody();     
             var csindex= casestr.search('noop">');
             caseid=casestr.substring(csindex+10,csindex+26);
             // caseid.push(casestr.substring(csindex-177,csindex+39));
             casedt=t.getDate();
             //msgid.push(t.getId());
             // Pulling my required information form the email body
             agent=casestr.substring(casestr.search('Hi <strong>')+11,casestr.search(',</strong>'));
             evelink=casestr.substring(casestr.search('https://google.com/'),casestr.search('https://google.com/')+77);
             workflow=casestr.substring(casestr.search('Work')+1610,casestr.indexOf('</p></div>',casestr.search('Work')+1610)); //added
             queue= casestr.substring(casestr.search('Queue')+1571,casestr.indexOf('</p></div>',casestr.search('Queue')+1571));   //added
             alignment= casestr.substring(casestr.search(' score')+1581,casestr.indexOf('</p></div>',casestr.search(' score')+1581)); //added
             t.star();
             if (threads[i].getFirstMessageSubject()=="email subject line")
                 {
                   appealdecision.appendRow([agent,caseid,casedt,link,work,queue,alignment]);
                 }
               else
               {appealactive.appendRow([agent,caseid,casedt,link,work,queue]);}
           }           
     }         
   }
  }
}
