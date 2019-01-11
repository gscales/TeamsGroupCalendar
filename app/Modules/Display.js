const buildScheduleTable = (Schedules,displayNameMap) => {
    var JSONData = [];   
    for (index = 0; index < Schedules.value.length; ++index) {
        var entry = Schedules.value[index].scheduleItems;
        entry.forEach(function (CalendarEntry) { 
            calEntry ={};
            calEntry.title = CalendarEntry.subject + " (" + displayNameMap[Schedules.value[index].scheduleId].initials + ")";            
            if(CalendarEntry.start.dateTime.slice(12,8) == "00:00:00"){ 
                calEntry.start = CalendarEntry.start.dateTime.slice(0,11);
                calEntry.end = CalendarEntry.end.dateTime.slice(0,11);    
            }else{
                calEntry.start = CalendarEntry.start.dateTime;
                calEntry.end = CalendarEntry.end.dateTime;    
            }        
            calEntry.color =  displayNameMap[Schedules.value[index].scheduleId].colorEntry;
            JSONData.push(calEntry);
        });  

    }
    return JSONData;
}

const buildLegend = (displayNameMap) => {
    var html = "<div class=\"ms-Table\" style=\"border-collapse:collapse;border: 0px;table-layout: auto;width:100%;\;background-color:white;\"><div class=\"ms-Table-row\">";
    html = html + "<span class=\"ms-Table-cell\" style=\"background-color:white;font-size: large;width:50px;font-weight:bolder;\"></span>";
    html = html + "<span class=\"ms-Table-cell\" style=\"background-color:white;font-size: large;width:150px;font-weight:bolder;\">Member</span>";
    html = html + "</div>";
    for (var key in displayNameMap) {
        if (displayNameMap.hasOwnProperty(key)) {
            console.log(key);
            console.log(displayNameMap[key].colorEntry);
            html = html + "<div class=\"ms-Table-row\"><span class=\"ms-Table-cell\" style=\"width:50px;\"><img id=\"img" + key + "\" style=\"border: 2px solid " + displayNameMap[key].colorEntry  + ";\" src=\"\" /></span>";
            html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\" style=\"vertical-align: middle;width:150px;background-color:" + displayNameMap[key].colorEntry + ";\">" + displayNameMap[key].displayName + "</span>";
            html = html + "</div >";
        }
    }
    html = html + "</div>";
    return html;
}