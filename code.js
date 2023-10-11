function getData(sheetName) { //Returns a rectangular grid of values in a given sheet.
    var data = SpreadsheetApp.getActive().getSheetByName(sheetName).getDataRange().getValues();
    return data;
  }
  
  function removeDuplicatesAndStore(){//Removes duplicates from source sheet,retains the first instance and pastes duplicates in Duplicates
    var sourceSheetName = "Copy"; // Name of the source sheet
    var destinationSheetName = "Duplicates"; // Name of the destination sheet
  
    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
    var values = sourceSheet.getDataRange().getValues();
  
    var columnIndex = 2; // Specify the column index (1-indexed) to check for duplicates
  
    var uniqueValues = [];
    var duplicateValues = [];
  
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var value = row[columnIndex - 1]; // Get the value from the specified column
      
      var isDuplicate = false;
      for (var j = 0; j < uniqueValues.length; j++) {
       
        if (uniqueValues[j][columnIndex - 1] === value) {
          duplicateValues.push(row);
          isDuplicate = true;
          break;
        }
      }
  
      if (!isDuplicate) {
        uniqueValues.push(row);
      }
    }
  
    var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var destinationSheet = destinationSpreadsheet.getSheetByName(destinationSheetName);
    if (!destinationSheet) {
      destinationSheet = destinationSpreadsheet.insertSheet(destinationSheetName);
      destinationSheet.getRange(1, 1, 1, values[0].length).setValues([values[0]]); // Copy headers to the destination sheet
    }
  
    if (duplicateValues.length > 0) {
      destinationSheet.getRange(destinationSheet.getLastRow() + 1, 1, duplicateValues.length, duplicateValues[0].length).setValues(duplicateValues);
    }
  
    sourceSheet.clearContents(); // Clear the contents of the source sheet
  
    if (uniqueValues.length > 0) {
      sourceSheet.getRange(1, 1, uniqueValues.length, uniqueValues[0].length).setValues(uniqueValues);
    }
  }
  
  
  
  function indicesOfLocationRequest(){ //returns of indices of rows of request in the spreadsheet for each location as an array
    indices=[[],[],[],[],[]];
    var values = getData("Copy");
    length=values.length;
    var locations = ["Mumbai","Hyderabad","Delhi","Bangalore","Chennai"]
    for(j=0;j<5;j++){
      for(i=0;i<length;i++){
        location = locations[j];
      if (values[i][6] == location){
        indices[j].push(i);
      }
      }
    } 
    
    return indices;
  }  
  
  function users(){ //data of users is stored
   var usersAtLocations= [];
   indices = indicesOfLocationRequest();
   data=getData("Copy");
   for(i=0;i<5;i++){
    usersAtLocations[i]=[];
    indicesAtLocation=indices[i];
    if(indicesAtLocation.length<=6){
      n=indicesAtLocation.length;
    }
    else{
      n=6;
    }
      for(k=0;k<n;k++){
      priority=[]
      for(j=7;j<13;j++){
      priority.push(data[indicesAtLocation[k]][j])
      }
      usersAtLocations[i].push({name:data[indicesAtLocation[k]][2],email:data[indicesAtLocation[k]][1],
      gender:data[indicesAtLocation[k]][5],preferences:priority });
    }
   
  }
    return usersAtLocations;
  }
  
  function allocateAppointments() { //appointments are allocated
   
    indices = indicesOfLocationRequest();
   
    
    usersAtLocation=users();
   
    slotsAtLocations=[];
     var timeSlots = [
      "10:00 AM - 10:30 AM","10:30 AM - 11:00 AM","11:00 AM - 11:30 AM","11:30 AM - 12:00 PM","12:00 PM - 12:30 PM","12:30 PM - 1:00 PM"];
     
    for(n=0;n<5;n++){
       var slots=[]
      
   var users2=usersAtLocation[n]
   
  
   appointmentSlotFilled=[0,0,0,0,0,0];
   appointmentUser=[0,0,0,0,0,0]
  
   for(i=0;i<users2.length;i++){
    for(j=1;j<=6;j++){
      k=users2[i].preferences.indexOf(j);
      if(appointmentSlotFilled[k]!==1){
       appointmentSlotFilled[k]=1;
       appointmentUser[k]=i+1;
       
     
       break;
      }
    }
   }
  
     
    // Initialize slots with time and allocatedUser properties
    for (var i = 0; i < timeSlots.length; i++) {
      slots.push({ time: timeSlots[i], allocatedUserEmail: null,allocatedUserName:null,allocatedUserGender:null});
    }
  
    for(i=1;i<=6;i++){
      l=appointmentUser.indexOf(i);
      if(l!==-1){
        slots[l].allocatedUserEmail=users2[i-1].email;
        slots[l].allocatedUserName=users2[i-1].name;
        slots[l].allocatedUserGender=users2[i-1].gender;
      }
    }
    slotsAtLocations.push(slots);
   
  
    // Print the allocation results in console
    
      for (var k = 0; k < slots.length; k++) {
      var slot = slots[k];
     
     
    }
    
    
    
  }
  
    
    return slotsAtLocations;
  }
  
  function waitingList(){ //indices of people in waiting list
    indices=indicesOfLocationRequest();
   
    waitingList=[[],[],[],[],[]];
    for(i=0;i<5;i++){
     
       list=[];
       
      if(indices[i].length>6){
       
         for(j=6;j<indices[i].length;j++){
          list.push(indices[i][j])
         }
        
         waitingList[i]=list;
      }
    }
   
   return waitingList
  }
  
  
  function sendEmails() { // mails are sent before cancellation is done
    var templateData = getData("Template");
    var emailSubject = templateData[1][0]; //Cell A2 (contains the email subject)
    slotsAtLocations = allocateAppointments();
     var locations = ["Mumbai","Hyderabad","Delhi","Bangalore","Chennai"]
    var emailData = getData("Copy");
    
    token_number=1
    for(i=0;i<5;i++){
    var location = locations[i];
  
  for(j=0;j<slotsAtLocations[i].length;j++){
   var user = slotsAtLocations[i][j].allocatedUserName;
   if(slotsAtLocations[i][j].allocatedUserGender=="Male"){
    user="Mr."+user;
   }
   var email= slotsAtLocations[i][j].allocatedUserEmail;
   var allocatedSlot = slotsAtLocations[i][j].time;
   var emailBody =  user + ",\n\n" +
                      "We are delighted to inform you that your appointment slot has been allocated!\n\n" +
                      "Appointment Details:\n" +
                      "- Time Slot: " +  allocatedSlot + "\n\n" +
                      "Token number: " + token_number+"\n"+
                      "Location"+ ":" + location + "\n\n"+
                      "Please make a note of this appointment and arrive on time. We look forward to see you!\n\n" +
                      "Best regards,\n" +
                      "Your Appointment Team"; 
  
    
    //var headerRow = emailData.shift(); //Remove the header row
  
    if(email!==null && user!==null){
      // Logger.log(email); "\n"
      // Logger.log(emailSubject);"\n"
      // Logger.log(emailBody); "\n"
      // "\n"
      token_number=token_number+1;
     MailApp.sendEmail(email, emailSubject, emailBody);
    }
    }}
  
  
  indices=indicesOfLocationRequest();
  
  for(i=0;i<indices.length;i++){
   
     
    if(indices[i].length>6){
     
      for(j=6;j<indices[i].length;j++){
        var waitingListNumber=j-5;
       if(emailData[indices[i][j]][5]=="Male"){
        gen="Mr."
       }
       else{
        gen="Ms."
       }
      
        var emailBody2= gen+emailData[indices[i][j]][2]+",\n\n"+
        
  
  "We regret to inform you that we were unable to schedule an appointment for you at this time." +"\n"+
   "We have received your request and unfortunately, all the available slots have been filled at the moment. However, you have been added to the waiting list.\n\n"+
  
  "Your waiting list number is: "+ waitingListNumber +"\n\n"+
  
  "We will notify you if a spot becomes available.\n\n"+
  "Best regards,\n"+
  "Your Appointment Team"; 
        email=emailData[indices[i][j]][1]
        emailSubject2= "Waiting List";
        // Logger.log(email)
        // Logger.log(emailSubject2)
        // Logger.log(emailBody2)
       MailApp.sendEmail(email, emailSubject2, emailBody2); 
      }
    }}
   // return token_number;
  
  
  emailData2=getData("Duplicates");
  
  
   emailData2.forEach(function (row) {
      var email = row[1];
     var emailSubject3= "Cancellation of Duplicate Appointment Request";
     if(row[5]=="Male"){
      gen2="Mr."
     }
     else{
      gen2="Ms."
     }
     var emailBody3= gen2+ row[2]+",\n\n"+
  
  "We hope this email finds you well. We would like to inform you that we have received your recent appointment request; however, we regret to inform you that your request has been canceled.\n\n"+
  
  "Upon reviewing our records, we have discovered that you already have a confirmed appointment slot scheduled. Our system detected the duplicate request.\n\n"+
  
  "Please be assured that your original appointment remains scheduled as planned, and there is no need for any further action.\n"
  "Thank you for your understanding and cooperation.\n\n"+
  
  "Best regards,\n"+
  "Your Appointment Team"; 
  if(email==" "){
    return;
  }
    //  Logger.log(email);"\n"
    //  Logger.log(emailSubject3);"\n";
    //  Logger.log(emailBody3);"\n";
  
      MailApp.sendEmail(email, emailSubject3, emailBody3);
    });
   return token_number;
  }
  
  function sendEmailsForCancellationsRequests(token){ //cancellation mails are sent and people who get appointment now get mails
    
    var token2=token;
   
   var emailData=getData("Cancellations");
   var emailData2=getData("Copy")
   var emailSubject2= "Updated Appointment Confirmation";
   var slotsAtLocations=allocateAppointments();
    
   var emailSubject= "Appointment Cancellation Confirmation";
    var headerRow = emailData.shift(); //Remove the header row
   var waiting = waitingList();
  
  
  emailData.forEach(function (row) {
   var email2=null;
  
   var username=row[2];
    if(row[5]=="Male"){
           username="Mr."+username
          }
          else{username="Ms."+username}
  
    var emailBody=  username+"\n\nWe have received your request to cancel your appointment . Your cancellation has been processed successfully.\n\n"+
  
  "We understand that circumstances may change, and we appreciate you notifying us in advance. If you would like to reschedule your appointment in the future, please don't hesitate to contact us.\n\n"+
  
  "If you have any questions or need further assistance, please feel free to reach out to our support team.\n"+
  
  "Thank you for your understanding.\n\n"+
  
  "Best regards,\n"+
  "Your Appointment Team"; 
  var n=-1;
  
      var email = row[1];
      // Logger.log(email)
      // Logger.log(emailSubject)
      // Logger.log(emailBody)
      MailApp.sendEmail(email, emailSubject, emailBody);
      
      if(row[6]=="Mumbai"){
        if(waiting[0].length!=0){
          num=waiting[0][0]
          name2=emailData2[num][2]
          if(emailData2[num][5]=="Male"){
            name2="Mr."+name2
          }
          else{name2="Ms."+name2}
          email2=emailData2[num][1]
          location="Mumbai";
           n=0;
          waiting[0].shift()
        }
        
      }
      else if(row[6]=="Hyderabad"){
        if(waiting[1].length!=0){
          num=waiting[1][0]
          name2=emailData2[num][2]
          if(emailData2[num][5]=="Male"){
            name2="Mr."+name2
          }
          else{name2="Ms."+name2}
          email2=emailData2[num][1]
          location="Hyderabad";
           n=1;
          token2++;
          waiting[1].shift()
        }
        
      }
      else if(row[6]=="Delhi"){
         if(waiting[2].length!=0){
          num=waiting[2][0]
          name2=emailData2[num][2]
           if(emailData2[num][5]=="Male"){
            name2="Mr."+name2
          }
          else{name2="Ms."+name2}
          email2=emailData2[num][1]
          location="Delhi";
           n=2;
           token2++;
          waiting[2].shift()
        }
      }
      else if(row[6]=="Bangalore"){
         if(waiting[3].length!=0){
          num=waiting[3][0]
          name2=emailData2[num][2]
           if(emailData2[num][5]=="Male"){
            name2="Mr."+name2
          }
          else{name2="Ms."+name2}
          if(emailData2[num][5]=="Male"){
            name="Mr."+name
          }
          else{name="Ms."+name}
          email2=emailData2[num][1]
          location="Bangalore";
           n=3;
           token2++;
          waiting[3].shift()
        }
      }
      else if(row[6]=="Chennai"){
         if(waiting[4].length!=0){
          num=waiting[4][0]
          name2=emailData2[num][2]
           if(emailData2[num][5]=="Male"){
            name2="Mr."+name2
          }
          else{name2="Ms."+name2}
          email2=emailData2[num][1]
          location="Chennai";
           n=4;
           token2++;
          waiting[4].shift()
        }
      }
      if(n!=-1){
      for(k=0;k<slotsAtLocations[n].length;k++){
        
        if(email==slotsAtLocations[n][k].allocatedUserEmail){
          allocatedSlot=slotsAtLocations[n][k].time;
        }
      }}
     
      
     
      
     var emailBody2 =  name2 + ",\n\n" +
    "We are delighted to inform you that a slot has become available, and your appointment has been scheduled!\n\n" +
    "Updated Appointment Details:\n" +
    "- Time Slot: " + allocatedSlot + "\n" +
    "Location: "+location+"\n"+
    "Token Number: "+token2+"\n\n"+
    "Please make a note of this updated appointment and arrive on time. We look forward to seeing you!\n\n" +
    "Best regards,\n" +
    "Your Appointment Team";
  
   if(email2!==null){
      //  Logger.log(email2)
      //  Logger.log(emailSubject2)
      //  Logger.log(emailBody2)   
      MailApp.sendEmail(email2, emailSubject2, emailBody2);
     }
    });
  
  
  }
  
  function appointmentToken(){
    //removeDuplicatesAndStore();
    token=sendEmails();
    sendEmailsForCancellationsRequests(token);
  }
