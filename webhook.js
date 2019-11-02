function doPost(e) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1juwIFGiCxwZvMqGXZVeZKY/edit");
  var sheet = ss.getSheetByName("แผ่น1");

  var data = JSON.parse(e.postData.contents)
  var userMsg = data.originalDetectIntentRequest.payload.data.message.text;
  var values = sheet.getRange(2, 2, sheet.getLastRow(),sheet.getLastColumn()).getValues(); //column1 [ID]
  var id = ["","","","","",""];
  var name = ["","","","","",""];
  var surname = ["","","","","",""];
  var nickname = ["","","","","",""];
  var department = ["","","","","",""];
  var extension = ["","","","","",""];
  var mobile = ["","","","","",""];
  var location = ["","","","","",""];
  var Count = 0; 

  for(var i = 0;i<values.length; i++){
    if(values[i][0] == userMsg || values[i][1] == userMsg || values[i][2] == userMsg || values[i][3] == userMsg){
      i=i+2;
    var founded = true;
    Count = Count+1; 

     id[Count] = sheet.getRange(i,1).getValue();
     name[Count] = sheet.getRange(i,2).getValue();
     surname[Count] = sheet.getRange(i,3).getValue();
     nickname[Count] = sheet.getRange(i,4).getValue();
     department[Count] = sheet.getRange(i,6).getValue();
     extension[Count] = sheet.getRange(i,7).getValue();
     mobile[Count] = sheet.getRange(i,8).getValue();    
     location[Count] = sheet.getRange(i,9).getValue();
  }
}

if(Count == 1){
      var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]
    }
   }
  }
]
}
}
else if(Count == 2){
        var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[2]+"\n"+
      "ชื่อ-นามสกุล :"+name[2]+" "+surname[2]+"\n"+
      "ชื่อเล่น : "+nickname[2]+"\n" +
      "ฝ่าย :"+department[2]+"\n"+
      "เบอร์โต๊ะ :"+extension[2]+"\n" +
      "เบอร์มือถือ :"+mobile[2]+"\n" +
      "สังกัด :"+location[2]
    }
   }
  }
]
}
}
else if(Count == 3){
var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[2]+"\n"+
      "ชื่อ-นามสกุล :"+name[2]+" "+surname[2]+"\n"+
      "ชื่อเล่น : "+nickname[2]+"\n" +
      "ฝ่าย :"+department[2]+"\n"+
      "เบอร์โต๊ะ :"+extension[2]+"\n" +
      "เบอร์มือถือ :"+mobile[2]+"\n" +
      "สังกัด :"+location[2]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[3]+"\n"+
      "ชื่อ-นามสกุล :"+name[3]+" "+surname[3]+"\n"+
      "ชื่อเล่น : "+nickname[3]+"\n" +
      "ฝ่าย :"+department[3]+"\n"+
      "เบอร์โต๊ะ :"+extension[3]+"\n" +
      "เบอร์มือถือ :"+mobile[3]+"\n" +
      "สังกัด :"+location[3]
    }
   }
  }
]
}
}
else if(Count == 4){
  var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[2]+"\n"+
      "ชื่อ-นามสกุล :"+name[2]+" "+surname[2]+"\n"+
      "ชื่อเล่น : "+nickname[2]+"\n" +
      "ฝ่าย :"+department[2]+"\n"+
      "เบอร์โต๊ะ :"+extension[2]+"\n" +
      "เบอร์มือถือ :"+mobile[2]+"\n" +
      "สังกัด :"+location[2]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[3]+"\n"+
      "ชื่อ-นามสกุล :"+name[3]+" "+surname[3]+"\n"+
      "ชื่อเล่น : "+nickname[3]+"\n" +
      "ฝ่าย :"+department[3]+"\n"+
      "เบอร์โต๊ะ :"+extension[3]+"\n" +
      "เบอร์มือถือ :"+mobile[3]+"\n" +
      "สังกัด :"+location[3]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[4]+"\n"+
      "ชื่อ-นามสกุล :"+name[4]+" "+surname[4]+"\n"+
      "ชื่อเล่น : "+nickname[4]+"\n" +
      "ฝ่าย :"+department[4]+"\n"+
      "เบอร์โต๊ะ :"+extension[4]+"\n" +
      "เบอร์มือถือ :"+mobile[4]+"\n" +
      "สังกัด :"+location[4]
    }
   }
  }
]
}
}
else if(Count == 5){
  var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[2]+"\n"+
      "ชื่อ-นามสกุล :"+name[2]+" "+surname[2]+"\n"+
      "ชื่อเล่น : "+nickname[2]+"\n" +
      "ฝ่าย :"+department[2]+"\n"+
      "เบอร์โต๊ะ :"+extension[2]+"\n" +
      "เบอร์มือถือ :"+mobile[2]+"\n" +
      "สังกัด :"+location[2]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[3]+"\n"+
      "ชื่อ-นามสกุล :"+name[3]+" "+surname[3]+"\n"+
      "ชื่อเล่น : "+nickname[3]+"\n" +
      "ฝ่าย :"+department[3]+"\n"+
      "เบอร์โต๊ะ :"+extension[3]+"\n" +
      "เบอร์มือถือ :"+mobile[3]+"\n" +
      "สังกัด :"+location[3]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[4]+"\n"+
      "ชื่อ-นามสกุล :"+name[4]+" "+surname[4]+"\n"+
      "ชื่อเล่น : "+nickname[4]+"\n" +
      "ฝ่าย :"+department[4]+"\n"+
      "เบอร์โต๊ะ :"+extension[4]+"\n" +
      "เบอร์มือถือ :"+mobile[4]+"\n" +
      "สังกัด :"+location[4]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[5]+"\n"+
      "ชื่อ-นามสกุล :"+name[5]+" "+surname[5]+"\n"+
      "ชื่อเล่น : "+nickname[5]+"\n" +
      "ฝ่าย :"+department[5]+"\n"+
      "เบอร์โต๊ะ :"+extension[5]+"\n" +
      "เบอร์มือถือ :"+mobile[5]+"\n" +
      "สังกัด :"+location[5]
    }
   }
  }
]
}
}
else if(Count == 6){
  var result = {
    "fulfillmentMessages": [
  {
    "platform": "line",
    "type": 4,
    "payload" : {
    "line":  {
    "type": "text",
      "text":
      "ค้นหา "+userMsg+"\n"+
      "รหัสพนักงาน "+id[1]+"\n"+
      "ชื่อ-นามสกุล :"+name[1]+" "+surname[1]+"\n"+
      "ชื่อเล่น : "+nickname[1]+"\n" +
      "ฝ่าย :"+department[1]+"\n"+
      "เบอร์โต๊ะ :"+extension[1]+"\n" +
      "เบอร์มือถือ :"+mobile[1]+"\n" +
      "สังกัด :"+location[1]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[2]+"\n"+
      "ชื่อ-นามสกุล :"+name[2]+" "+surname[2]+"\n"+
      "ชื่อเล่น : "+nickname[2]+"\n" +
      "ฝ่าย :"+department[2]+"\n"+
      "เบอร์โต๊ะ :"+extension[2]+"\n" +
      "เบอร์มือถือ :"+mobile[2]+"\n" +
      "สังกัด :"+location[2]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[3]+"\n"+
      "ชื่อ-นามสกุล :"+name[3]+" "+surname[3]+"\n"+
      "ชื่อเล่น : "+nickname[3]+"\n" +
      "ฝ่าย :"+department[3]+"\n"+
      "เบอร์โต๊ะ :"+extension[3]+"\n" +
      "เบอร์มือถือ :"+mobile[3]+"\n" +
      "สังกัด :"+location[3]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[4]+"\n"+
      "ชื่อ-นามสกุล :"+name[4]+" "+surname[4]+"\n"+
      "ชื่อเล่น : "+nickname[4]+"\n" +
      "ฝ่าย :"+department[4]+"\n"+
      "เบอร์โต๊ะ :"+extension[4]+"\n" +
      "เบอร์มือถือ :"+mobile[4]+"\n" +
      "สังกัด :"+location[4]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[5]+"\n"+
      "ชื่อ-นามสกุล :"+name[5]+" "+surname[5]+"\n"+
      "ชื่อเล่น : "+nickname[5]+"\n" +
      "ฝ่าย :"+department[5]+"\n"+
      "เบอร์โต๊ะ :"+extension[5]+"\n" +
      "เบอร์มือถือ :"+mobile[5]+"\n" +
      "สังกัด :"+location[5]+"\n"
      +"\n"+
      "รหัสพนักงาน "+id[6]+"\n"+
      "ชื่อ-นามสกุล :"+name[6]+" "+surname[6]+"\n"+
      "ชื่อเล่น : "+nickname[6]+"\n" +
      "ฝ่าย :"+department[6]+"\n"+
      "เบอร์โต๊ะ :"+extension[6]+"\n" +
      "เบอร์มือถือ :"+mobile[6]+"\n" +
      "สังกัด :"+location[6]
    }
   }
  }
]
}
}
    var replyJSON = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    return replyJSON;
  }
