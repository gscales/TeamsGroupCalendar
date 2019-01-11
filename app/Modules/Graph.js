const GetGroupMembers = (idToken, teamscontext) => {
    return new Promise(
        (resolve, reject) => {
            GroupId = teamscontext.groupId;
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                url: ("https://graph.microsoft.com/v1.0/groups/" + GroupId + "/members"),
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + idToken }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}
const GetSchedule = (idToken, GroupMembers, displayNameMap) => {
    return new Promise(
        (resolve, reject) => {
            var SchPost = {};
            SchPost.schedules = [];
            for (index = 0; index < GroupMembers.value.length; ++index) {
                var entry = GroupMembers.value[index];
                SchPost.schedules.push(entry.mail);
                var dnMapValue = {};
                dnMapValue.displayName = entry.displayName;
                var initials = "";
                if(entry.givenName.length > 0){
                    initials = entry.givenName.slice(0,1);
                }
                if(entry.surname.length > 0){
                    initials = initials + entry.surname.slice(0,1);
                }
                dnMapValue.initials = initials;
                dnMapValue.colorEntry =  randomColor({luminosity: 'bright',format: 'hsla'});            
                displayNameMap[entry.mail] = dnMapValue;
            }
            var Start = new Date();
            var End = new Date();
            End.setDate(End.getDate() + 62); 
            var StartTime = {};
            StartTime.TimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            StartTime.dateTime = formatDate(Start) + "T08:00:00";
            var EndTime = {};
            EndTime.dateTime = formatDate(End) + "T08:00:00";
            EndTime.TimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            SchPost.startTime = StartTime;
            SchPost.endTime = EndTime;
            SchPost.availabilityViewInterval = 1440;
            var schRequest = JSON.stringify(SchPost);
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: ("https://graph.microsoft.com/beta/me/calendar/getSchedule"),
                dataType: 'json',
                data: schRequest,
                headers: {
                    'Authorization': 'Bearer ' + idToken,
                    'Prefer': ('outlook.timezone="' + Intl.DateTimeFormat().resolvedOptions().timeZone + "\""),

                }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}

const GetUserPhotos = (idToken,GroupMembers) => {
    var photoRequestMap = {};
    for (index = 0; index < GroupMembers.value.length; ++index) {
        var entry = GroupMembers.value[index];
        var clientid = uuidv4();
        photoRequestMap[clientid] = entry.mail;
        var userImageURL = "https://graph.microsoft.com/v1.0/users('" + entry.mail + "')/photos/48x48/$value";
        var xhr = new XMLHttpRequest();
        xhr.open('GET', userImageURL, true);
        xhr.setRequestHeader("Authorization", "Bearer " + idToken);
        xhr.setRequestHeader("client-request-id", clientid);
        xhr.responseType = 'arraybuffer';
        xhr.onload = function (e) {
            if (this.status == 200) {
                // get binary data as a response
                var blob = this.response;
                var clientRespHeader = this.getResponseHeader("client-request-id")
                var ElemendId = "img" + photoRequestMap[clientRespHeader];
                var uInt8Array = new Uint8Array(this.response);
                var data = String.fromCharCode.apply(String, uInt8Array);
                var base64 = window.btoa(data);
                document.getElementById(ElemendId).src = "data:image/png;base64," + base64;

            }
        };

        xhr.send()
        $('#ShowBoard').hide();
        $('#ProgresLoader').hide();
    }
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}







