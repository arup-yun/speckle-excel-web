'use strict';
var ws = {};
//var apiAddress = 'https://ireland-test.speckle.arup.com/api';
//Arup 的是指 https://ireland-test.speckle.arup.com/api
//hestia 指的是 https://hestia.speckle.works/api/
Office.onReady(function () {
    // Office is ready
    $(document).ready(function () {
        init();
        $("#streams-tab").click(function () {
            $('#accounts-content').hide();
            $('#streams-content').show();
            $("#streams-tab").addClass('active');
            $("#accounts-tab").removeClass('active');
        });
        $("#accounts-tab").click(function () {
            $('#accounts-content').show();
            $('#streams-content').hide();
            $("#streams-tab").removeClass('active');
            $("#accounts-tab").addClass('active');
        });
        $(".list-group,.listgroupA").click(function (event) {
            event.stopPropagation();
            return false;
        });
    });
})

function init() {
    var myModalMsg = $('#myModalMsg');
    myModalMsg.html('');
    $('#doLoading').hide();
    $('#accounts-content').show();
    $('#streams-content').hide();
    getAccountList();
    getStreamsList();
}

function showMsg(type, str, dom,status) {
    dom = dom || '#myModalMsg';
    console.log(dom);
    var myModalMsg = $(dom);
    myModalMsg.html('');
    myModalMsg.html('<div class="alert alert-' + type + ' alert-dismissible" role="alert"><button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button><strong>' + str + '</strong ></div>');
    if (!status) {
        setTimeout(function () {
            myModalMsg.html('');
        }, 10000);
    }
    
}

function doOpenLogin() {
    $('#myModal').modal('show');
}

function validateLoginForm() {
    var apiUrl = $('#apiUrl');
    var username = $('#username');
    var password = $('#password');
    if (!apiUrl.val()) {
        showMsg('danger', 'Spcckle server api url Can\'t be empty', '#validateMsg',true);
        apiUrl.focus();
        return false;
    } else {

    }
    if (!username.val()) {
        showMsg('danger', 'Your email address Can\'t be empty', '#validateMsg', true);
        username.focus();
        return false;
    }
    if (!password.val()) {
        showMsg('danger', 'Your account password Can\'t be empty', '#validateMsg', true);
        password.focus();
        return false;
    }
    $('#validateMsg').html('');
    return true;
}

function userLogin() {
    var apiUrl = $('#apiUrl').val();
    var username = $('#username').val();
    var password = $('#password').val();
    if (validateLoginForm() == false) {
        return;
    } else {
        $('#myModal').modal('hide');
        $('#doLoading').show();
        $.ajax({
            url: apiUrl + '/accounts/login',
            contentType: "application/json",
            method: 'post',
            data: JSON.stringify({
                "email": username,
                "password": password
            }),
            success: function (res) {
                if (res.success === true) {
                    var apiToken = res.resource.apitoken;
                    getServerName({
                        apiUrl: apiUrl,
                        apiToken: apiToken,
                        email: username,
                        password: password
                    });
                    showMsg('success', 'Login successful');
                } else {
                    showMsg('warning', 'Login failed, please try again later');
                }
                $('#doLoading').hide();
            },
            error: function (e) {
                $('#doLoading').hide();
                showMsg('warning', 'Unknown error login failed, please try again later');
                console.log(e);
            }
        });
    }
    

}

function getServerName(obj) {
    $.ajax({
        url: obj.apiUrl,
        contentType: "application/json; charset=utf-8",
        method: 'get',
        success: function (res) {
            var serverName = res.serverName;
            var accountList = JSON.parse(getLocalStorage('accountList'));
            accountList[serverName + "," + obj.email] = {
                apiUrl: obj.apiUrl,
                apiToken: obj.apiToken,
                email: obj.email,
                password: obj.password,
                serverName: serverName
            };
            localStorage.setItem("accountList", JSON.stringify(accountList));
            getAccountList();
        },
        error: function (e) {
            console.log(e);
        }
    });
}

function deleteAccount(serverName) {
    var accountList = JSON.parse(getLocalStorage('accountList'));
    if (accountList[serverName]) {
        delete accountList[serverName];
        localStorage.setItem('accountList', JSON.stringify(accountList));
        getAccountList();
    }
}

function showDetail(dom) {
    var firstNode = $(dom).children(".iconDown");
    var secondNode = $(dom).children(".iconUp");
    var lastNode = $(dom).children(".list-group");
    if (lastNode.is(':hidden')) {
        lastNode.removeClass("hide");
        lastNode.show();

        secondNode.removeClass("hide");
        secondNode.show();
        firstNode.hide();
        
    } else {
        lastNode.hide();
        firstNode.show();
        secondNode.hide();
    }
}

function copyApiUrl(dom, msgDom, msg) {
    msg = msg || 'Address copy Successful';
    var rng = document.body.createTextRange();
    rng.moveToElementText(dom);
    rng.scrollIntoView();
    rng.select();
    rng.execCommand("Copy");
    rng.collapse(false);
    showMsg('success', msg , msgDom);
}

function getLocalStorage(str) {
    if (localStorage.getItem(str) != null) {
        return localStorage.getItem(str);
    }
    return '{}';
}

function getSessionStorage(str) {
    if (sessionStorage.getItem(str) != null) {
        return sessionStorage.getItem(str);
    }
    return '{}';
}

function getAccountList() {
    var accountsDom = $('#accounts-list-group');
    var domList = "";
    var accountList = JSON.parse(getLocalStorage('accountList'));
    console.log(accountList);
    if (accountList !== {}) {        
        for (var k in accountList) {
            domList = domList + '<a href="#" class="list-group-item" onclick="showDetail(this)">' +
                '<span class="glyphicon glyphicon-menu-down icon iconDown pointer"></span>' +
                '<span class="glyphicon glyphicon-menu-up icon iconUp hide pointer"></span>' +
                '<h3>' + accountList[k].serverName + '</h3>' +
                '<p>' + accountList[k].email + '</p>' +
                '<ul class="list-group hide listgroupA">' +
                '<li class="list-group-item">' +
                '<span class="glyphicon glyphicon-envelope"></span>' +
                accountList[k].email + '</li>' +
                '<li class="list-group-item">' +
                '<span class="glyphicon glyphicon-check"></span>' +
                '<span onclick="copyApiUrl(this)" class="pointer">' + accountList[k].apiUrl + '</span>' +
                '<p class="textRigth mt10">' +
                '<button type="button" class="btn btn-danger btn-sm" onclick="deleteAccount(\'' + accountList[k].serverName + ',' + accountList[k].email + '\')">Delete</button>' +
                '</p></li></ul></a>';
        }
        accountsDom.html(domList);
    }
}

function ExcelError(err) {
    console.log(err);
}

function getStreamsList() {
    var streamsList = JSON.parse(getLocalStorage('streamsList'));
    var bindStreamsObj = JSON.parse(getLocalStorage('bindStreamsObj'));
    var excelBindStreamsObj = {};
    Excel.run(function (ctx) {
        var settings = ctx.workbook.settings;
        var loadObj = {};
        settings.load('items');
        return ctx.sync().then(function () {
            var setList = settings.items;
            var streamsDom = $('#streamsList');
            setList.forEach(function (item) {
                if (item.m_value && bindStreamsObj[item._K] && streamsList[item._K]) {
                    excelBindStreamsObj[item._K] = true;
                    var rowDom = '<a href="#" id="streamId-' + item._K + '" class="list-group-item ItemDeleteShow" style="border:none;" onclick="closeBtnGroup(\''+ item._K +'\')">' +
                        '<ul class="list-group">' +
                        '<li class="list-group-item classAbsolute">' +
                        '<div class="Mask-loading" id="itemLoading-' + item._K + '">' +
                        '<div class="loadingDiv" >' +
                        '<div class="k-line k-line11-1"></div> <div class="k-line k-line11-2"></div> <div class="k-line k-line11-3"></div >' +
                        '<div class="k-line k-line11-4"></div><div class="k-line k-line11-5"></div>' +
                        '</div></div >' +
                        '<span class="glyphicon glyphicon-list-alt mr10"></span>' +
                        '<span>' + streamsList[item._K].name +'</span>' +
                        '<p>' +
                        '<span class="streamsBtn pointer" onclick="copyApiUrl(this,\'#myAccountsMsg\',\'stream id copy Successful\')">' + item._K + '</span>' +
                        '<span class="textGray">Updated </span> <span id="' + item._K + '" class="textGray"> ' + checkUpdatedMinutes('#' + item._K, streamsList[item._K].updatedAt) + '</span><span  class="textGray"> minutes ago</span></p>' +
                        '<div class="iconBtnGroup">' +
                        '<div class="iconHover">' +
                        '<span class="glyphicon glyphicon-option-vertical iconPadding pointer" onclick="showBtnGroup(\'' + item._K +'\')" ></span > ' +
                        '<div class="btn-group-vertical btnGroup" role="group" id="' + item._K + 'BtnGroup">' +
                        '<button type="button" class="btn btn-default" onclick="getViewData(\'' + item._K + '\',\'' + streamsList[item._K].tableName + '\',\'' + streamsList[item._K].accountObj.apiUrl+'\')">View data</button>' +
                        '<button type="button" class="btn btn-default" onclick="getViewObjects(\'' + item._K + '\',\'' + streamsList[item._K].tableName + '\',\'' + streamsList[item._K].accountObj.apiUrl +'\')">View objects</button>' +
            '<button type="button" class="btn btn-default" onclick="deleteStreams(\'' + item._K + '\',\'' + streamsList[item._K].tableName + '\')">Delete</button>' +
                        '<button type="button" class="btn btn-default" onclick="zoomToTable(\'' + item._K + '\',\'' + streamsList[item._K].tableName + '\')">Zoom to table</button>' +
                        '</div></div></div ></li></ul></a>';
                    streamsDom.append(rowDom);
                    watchTablesChange(streamsList[item._K].tableName, item._K);
                }
            });
            sessionStorage.setItem('excelBindStreamsObj', JSON.stringify(excelBindStreamsObj));
        });
    }).catch(ExcelError);
}

function zoomToTable(streamId, tableName) {
    $('#' + streamId + 'BtnGroup').hide();
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.getItem(tableName);
        var tableRange = table.getRange();
        tableRange.select();
        return ctx.sync();
    }).catch(ExcelError);
}

function doSender() {
    var getAcount = getLocalStorage('accountList');
    if (getAcount === '{}') {
        showMsg('danger', 'please  Login first', '#myAccountsMsg');
        return;
    }
    Excel.run(function (ctx) {
        var workbook = ctx.workbook;
        workbook.load('tables');
        return ctx.sync().then(function () {
            var tablesItem = workbook.tables.items;
            if (!tablesItem || tablesItem.length === 0) {
                showMsg('danger', 'Please insert the table first', '#myAccountsMsg',true);
                return;
            }

            var accountList = JSON.parse(getLocalStorage('accountList'));
            //add acount list
            var myAccount = $('#myAccount');
            var accountDom = '';
            for (var k in accountList) {
                var tempData = JSON.stringify(accountList[k]);
                accountDom = accountDom + '<li onclick="checkAccountVal(' + JSON.stringify(accountList[k]).replace(/\"/g, "'") +')"><a href="#">' + k + '</a></li>'; 
            }
            myAccount.html(accountDom);
            //add table list
            var myTables = $('#myTables');
            var tablesDom = '';
            for (var i = 0; i < tablesItem.length; i++) {
                tablesDom = tablesDom + '<li onclick="checkTableVal(\'' + tablesItem[i].name + '\')"><a href="#">' + tablesItem[i].name + '</a></li>';
            }
            myTables.html(tablesDom);
            doOpenSender();
        });
    }).catch(ExcelError);
}

function checkTableVal(str) {
    var tableName = $('#tableName');
    tableName.css('color', '#000');
    tableName.html(str);
    sessionStorage.setItem('tableName', str);
    $('#streamName').focus().val(str);
}

function checkAccountVal(obj) {
    var accountName = $('#accountName');
    accountName.css('color', '#000');
    accountName.html(obj.serverName + ',' + obj.email);
    sessionStorage.setItem('accountName', obj.serverName + ',' + obj.email);
    sessionStorage.setItem('accountObj', JSON.stringify(obj));
}

function doOpenSender() {
    $('#mySender').modal('show');
    $('#tableName').html('Please select the table');
    $('#tableName').css('color', '#757575');
    $('#streamName').val('');
    $('#accountName').html('Please select the Account');
    $('#accountName').css('color', '#757575');
    sessionStorage.setItem('tableName', '');
    sessionStorage.setItem('accountName','');
}

function validateAddForm() {
    var tableName = sessionStorage.getItem('tableName');
    var accountName = sessionStorage.getItem('accountName');
    
    if (!accountName) {
        showMsg('danger', 'Your Account Can\'t be empty', '#validateSenderMsg', true);
        password.focus();
        return false;
    }
    if (!tableName) {
        showMsg('danger', 'Your Excel table name Can\'t be empty', '#validateSenderMsg', true);
        username.focus();
        return false;
    }
    $('#validateSenderMsg').html('');
    return true;
}

function getTableData(tableName, streamName) {
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.getItem(tableName);
        // Get data from the header row
        var headerRange = table.getHeaderRowRange().load("values");

        // Get data from the table
        var bodyRange = table.getDataBodyRange().load("values");

        ctx.sync().then(function () {
            var values = bodyRange.values; 
            var objects = [];
            var tableTitle = headerRange.values[0];
            sessionStorage.setItem(tableName + 'Count', values.length);

            for (let i = 0; i < values.length; i++) {
                var row = values[i];
                var object = {};
                for (var j = 0; j < tableTitle.length; j++) {
                    object[tableTitle[j]] = row[j];
                }
                objects.push({
                    type: "Object",
                    properties: object,
                    applicationId : 'object' + i + Math.random()
                });
            }
            createdObjects(objects, streamName);
        }); 
        
    }).catch(ExcelError);
}

function createdObjects(objects, streamName, streamObj) {
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    console.log(accountObj.apiToken);
    console.log(JSON.stringify(objects));
    $.ajax({
        url: accountObj.apiUrl + '/objects',
        contentType: "application/json; charset=utf-8",
        method: 'post',
        async: false,
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "User-Agent": "PostmanRuntime/7.15.0",
            "cache-control": "no-cache"
        },
        processData: false,
        data: JSON.stringify(objects),
        success: function (res) {
            if (res.success) {
                var tableName = sessionStorage.getItem('tableName');
                localStorage.setItem(tableName + 'Resource', JSON.stringify(res.resources));
                if (streamObj) {
                    uploadStreams(res.resources, streamObj);
                    showMsg('success', 'Add objects has been created successfully', '#myAccountsMsg');
                } else {
                    createdStreams(streamName);
                    showMsg('success', 'Objects has been created successfully', '#myAccountsMsg');
                }
            } else {
                if (streamObj) {
                    showMsg('warning', 'Unknown error created objects failed, please try again later', '#myAccountsMsg');
                }
                
            }
            $('#doLoading').hide();
        },
        error: function (e) {
            $('#doLoading').hide();
            if (!streamObj) {
                showMsg('warning', 'Unknown error created objects failed, please try again later', '#myAccountsMsg');
            }
            console.log(e);
        }
    });
}

function createdStreams(streamName) {
    var tableName = sessionStorage.getItem('tableName');
    var objects = getLocalStorage(tableName + 'Resource');
    var objCount = sessionStorage.getItem(tableName + 'Count');
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    var query = {
        name: streamName || sessionStorage.getItem('tableName'),
        objects: JSON.parse(objects || '{}'),
        layers: [{
            name: streamName || tableName,
            orderIndex: 0,
            startIndex: 0,
            objectCount: objCount,
            topology: "0;0-" + objCount,
            guid: getGUID()
        }]
    }
    $.ajax({
        url: accountObj.apiUrl + '/streams',
        contentType: "application/json; charset=utf-8",
        method: 'post',
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "Cache-Control": "no-cache"
        },
        processData: false,
        data: JSON.stringify(query),
        success: function (res) {
            if (res.success) {
                bindingTableAndStreams(tableName, res.resource);
                getClientsId(res.resource.streamId, accountObj);
                showStreamsList(res.resource, tableName, accountObj);
                showMsg('success', 'Streams has been created successfully', '#myAccountsMsg');
            } else {
                showMsg('warning', 'Unknown error created streams failed, please try again later', '#myAccountsMsg');
            }
            $('#doLoading').hide();
            watchTablesChange(tableName,res.resource.streamId);
        },
        error: function (e) {
            $('#doLoading').hide();
            showMsg('warning', 'Unknown error created streams failed, please try again later', '#myAccountsMsg');
            console.log(e);
        }
    });
}

function getClientsId(streamId, accountObj) {
    var temp = {
        owner: accountObj.email,
        private: true,
        anonymousComments: true,
        deleted: false,
        streamId: streamId,
        online: true
    };
    $.ajax({
        url: accountObj.apiUrl + '/clients',
        contentType: "application/json; charset=utf-8",
        method: 'post',
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "Cache-Control": "no-cache"
        },
        processData: false,
        data: JSON.stringify(temp),
        success: function (res) {
            if (res.success) {
                var streamsList = JSON.parse(getLocalStorage('streamsList'));
                streamsList[streamId].clientId = res.resource._id;
                localStorage.setItem('streamsList', JSON.stringify(streamsList));
                BroadcastMessage(streamId);
                showMsg('success', 'Streams has been created successfully', '#myAccountsMsg');
            } else {
                showMsg('warning', 'Unknown error created streams failed, please try again later', '#myAccountsMsg');
            }
        },
        error: function (e) {
            $('#doLoading').hide();
            showMsg('warning', 'Unknown error created streams failed, please try again later', '#myAccountsMsg');
            console.log(e);
        }
    });
}

function bindingTableAndStreams(tableName,resource) {
    Excel.run(function (ctx) {
        var bindStreamsObj = JSON.parse(getLocalStorage('bindStreamsObj'));

        var settings = ctx.workbook.settings;
        settings.add(resource.streamId, true);

        bindStreamsObj[resource.streamId] = true;
        localStorage.setItem('bindStreamsObj', JSON.stringify(bindStreamsObj));


        var needsReview = settings.getItem(resource.streamId);
        needsReview.load("value");

        return ctx.sync().then(function () {
            console.log("Workbook needs review : " + needsReview.value);
        });
    }).catch(ExcelError);
}

//function getTableAndStreams(tableName) {
//    Excel.run(function (context) {
//        var customProperty = context.workbook.properties.custom;
//        var customPropertyCount = customProperty.getCount();

//        var customPropertys = customProperty.load("items");
       
//        //customProperty.load('items');
//        return context.sync(function () {
//            if (customPropertyCount.value > 0) {
//                customPropertys.forEach(function(prop){
//                    console.log(prop)
//                  });
//            } else {
//                console.log("No custom properties");
//            }
//        });
//    }).catch(ExcelError);
//}

function doAdd() {
    var tableName = sessionStorage.getItem('tableName');
    // validate that streams have been created
    var tableNameResource = getLocalStorage(tableName + 'Resource');
    if (tableNameResource != null) {
        showMsg('warning', 'The selected table has been created for streams. Please modify it and try again', '#validateSenderMsg');
    }

    var accountName = sessionStorage.getItem('accountName');
    var streamName = $('#streamName').val();
    if (validateAddForm()) {
        getTableData(tableName, streamName);
        $('#mySender').modal('hide');
        var accountObj = JSON.parse(getSessionStorage('accountObj'));
        $('#doLoading').show();
    }
    return;
}

function getGUID() {
    var s = [];
    var hexDigits = "0123456789abcdef";
    for (var i = 0; i < 36; i++) {
        s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
    }
    s[14] = "4"; // bits 12-15 of the time_hi_and_version field to 0010
    s[19] = hexDigits.substr((s[19] & 0x3) | 0x8, 1); // bits 6-7 of the clock_seq_hi_and_reserved to 01
    s[8] = s[13] = s[18] = s[23] = "-";

    var uuid = s.join("");
    return uuid;
}

function showStreamsList(streamsObj, tableName, accountObj) {
    var streamId = streamsObj.streamId;
    var streamsList = JSON.parse(getLocalStorage('streamsList'));
    var localNum = 0;
    if (streamsList[streamId]) {
        localNum++;
        streamsList[streamId] = streamsObj;
    }
    if (localNum > 0) {
        return;
    } else {
        streamsList[streamId] = streamsObj;
        streamsList[streamId].tableName = tableName;
        streamsList[streamId].accountObj = accountObj;
        var oldDom = $('#streamsList');
        var rowDom = '<a href="#" id="streamId-' + streamsObj.streamId+'" class="list-group-item ItemDeleteShow" style="border:none;">' +
            '<ul class="list-group">' +
            '<li class="list-group-item classRelative">' +
            '<div class="Mask-loading" id="itemLoading-' + streamsObj.streamId +'">'+
            '<div class="loadingDiv" >' +
            '<div class="k-line k-line11-1"></div> <div class="k-line k-line11-2"></div> <div class="k-line k-line11-3"></div >' +
            '<div class="k-line k-line11-4"></div><div class="k-line k-line11-5"></div>' +
            '</div></div >' +
            '<span class="glyphicon glyphicon-list-alt mr10"></span>' +
            '<span>' + streamsObj.name + '</span>' +
            '<p>' +
            '<span class="streamsBtn pointer" onclick="copyApiUrl(this,\'#myAccountsMsg\',\'stream id copy Successful\')">' + streamsObj.streamId + '</span>' +
            '<span class="textGray">Updated </span> <span id="' + streamsObj.streamId + '" class="textGray"> ' + checkUpdatedMinutes('#' + streamsObj.streamId, streamsObj.updatedAt) + '</span><span class="textGray"> minutes ago</span></p>' +
            '<div class="iconBtnGroup">' +
            '<div class="iconHover">' +
            '<span class="glyphicon glyphicon-option-vertical iconPadding pointer" onclick="showBtnGroup(\'' + streamsObj.streamId+'\')"></span>' +
            '<div class="btn-group-vertical btnGroup" role="group" id="' + streamsObj.streamId + 'BtnGroup">' +
            '<button type="button" class="btn btn-default" onclick="getViewData(\'' + streamsObj.streamId + '\',\'' + tableName + '\',\'' + accountObj.apiUrl + '\')">View data</button>' +
            '<button type="button" class="btn btn-default" onclick="getViewObjects(\'' + streamsObj.streamId + '\',\'' + tableName + '\',\'' + accountObj.apiUrl + '\')">View objects</button>' +
            '<button type="button" class="btn btn-default" onclick="deleteStreams(\'' + streamsObj.streamId + '\',\'' + tableName + '\')">Delete</button>' +
            '<button type="button" class="btn btn-default" onclick="zoomToTable(\'' + streamsObj.streamId + '\',\'' + tableName + '\')">Zoom to table</button>' +
            '</div></div></div></li></ul></a> ';
        oldDom.append(rowDom);
        
    }
    localStorage.setItem('streamsList', JSON.stringify(streamsList));
}

function showBtnGroup(streamId) {
    var e = window.event || arguments.callee.caller.arguments[0];
    e.stopPropagation();
    var streamBtnGroup = $('#' + streamId + 'BtnGroup');
    streamBtnGroup.css({ "display": "inline-block" });
    $('body').click(function () {
        streamBtnGroup.hide();
    })
}
function closeBtnGroup(streamId) {
    $('#' + streamId + 'BtnGroup').hide();
}

// Refresh time
function checkUpdatedMinutes(dom, date) {
    var now = new Date().getTime();
    var updated = new Date(date).getTime();
    var timeDiff = (updated - now) / 1000 / 60 ;
    var diffTime = Math.round(timeDiff);
    setTimeout(function() {
        updateStreamsTimer(dom, diffTime);
    },1000)
    return diffTime;
}

function updateStreamsTimer(dom, diffTime) {
    var domTimerObj = JSON.parse(getSessionStorage('domTimerObj'));
    if (domTimerObj[dom]) {
        clearInterval(domTimerObj[dom]);
    }
    domTimerObj[dom] = setInterval(function () {

        diffTime = diffTime + 1;
        $(dom).html(diffTime + '');
    }, 60000);
    sessionStorage.setItem('domTimerObj', JSON.stringify(domTimerObj));
}
function deleteStreams(streamId, tableName) {
    $('#ConfirmModal').modal('show');
    $('#confirm-content').html('Are you sure you want to delete the  table name " '+ tableName + ' "');
    $('#confirm-streamId').val(streamId);
}

function getViewData(streamId,tableName,url) {
    window.location.href = url + "/streams/" + streamId;
    $('#' + streamId + 'BtnGroup').hide();
}

function getViewObjects(streamId, tableName,url) {
    window.location.href = url + "/streams/" + streamId + '/objects/?omit=displayValue,base64';
    $('#' + streamId + 'BtnGroup').hide();
}

function doDeleteStreams() {
    $('#ConfirmModal').modal('hide');
    var streamId = $('#confirm-streamId').val();
    var streamsList = JSON.parse(getLocalStorage('streamsList'));
    if (streamsList[streamId]) {
        var tableName = streamsList[streamId].tableName;
        // delete UI
        $('#streamId-' + streamId).remove();

        // delete list
        delete streamsList[streamId];
        localStorage.setItem("streamsList", JSON.stringify(streamsList));

        // delete streams bindings
        var bindStreamsObj = JSON.parse(getLocalStorage('bindStreamsObj'));
        delete bindStreamsObj[streamId];
        localStorage.setItem("bindStreamsObj", JSON.stringify(bindStreamsObj));

        // delete timer
        var domTimerObj = JSON.parse(getSessionStorage('domTimerObj'));
        if (domTimerObj['#'+streamId]) {
            clearInterval(domTimerObj[streamId]);
        }
        delete domTimerObj['#'+streamId];
        sessionStorage.setItem('domTimerObj', JSON.stringify(domTimerObj));

        // delete excel table and bindings
        Excel.run(function (ctx) {
            var table = ctx.workbook.tables.getItem(tableName);
            var settings = ctx.workbook.settings;
            settings.add(streamId, false);
            var needsReview = settings.getItem(streamId);
            needsReview.load('value');

            // delete table
            //table.delete();
           
            return ctx.sync().then(function () {
                console.log("Workbook status : " + needsReview.value);
            });
        }).catch(ExcelError);
    }
    
}

function watchTablesChange(tableName, streamId) {
    var excelBindStreamsObj = JSON.parse(getSessionStorage('excelBindStreamsObj'));

    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.getItem(tableName);
        //table.onSelectionChanged.add(onSelectionChange);
        excelBindStreamsObj[streamId] = ctx.workbook.bindings.add(table.getRange(), "Table", streamId);
        excelBindStreamsObj[streamId].onDataChanged.add(onBindingDataChanged);
        return ctx.sync();
    }).catch(ExcelError);
}
//function onSelectionChange(eventArgs) {
//    console.log(eventArgs);
//}

function showStreamsLoading(streamId) {
    $('#itemLoading-' + streamId).show();
    console.log($('#itemLoading-' + streamId));
    console.log('show');
}

function hideStreamsLoading(streamId) {
    $('#itemLoading-' + streamId).hide();
}

function getStreamsDetail(streamId, headerRangeArr, bodyRangeArr) {
    if (!streamId) {
        return;
    }
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    $.ajax({
        url: accountObj.apiUrl + '/streams/' + streamId,
        contentType: "application/json; charset=utf-8",
        method: 'get',
        async: false,
        success: function (res) {
            if (res.success === true) {
                var objects = res.resource && res.resource.objects || [];
                sessionStorage.setItem('tableName', res.resource.name);
                checkBindStreamsObjAccount(streamId);
                checkObjects(streamId,objects,headerRangeArr[0], bodyRangeArr);
            } else {
                showMsg('warning', 'Get streams detail failed, please try again later');
            }
        },
        error: function (e) {
            showMsg('warning', 'Get streams detail failed, please try again later');
            hideStreamsLoading(streamId);
            console.log(e);
        }
    });
}

/* table uploaded auto upload objects
 * params: 
 * queryObjects is the upload streams's obects data
 * headerRangeArr is the object columns
 * bodyRangeArr is table data
 */
function uploadObjects(streamId, queryObjects, headerRangeArr, bodyRangeArr) {
    sessionStorage.setItem('uploadNum', queryObjects.length);
    queryObjects.forEach(function (item, index) {
        editObject(item._id, streamId, headerRangeArr, bodyRangeArr[index])
    });

}

function editObject(objectId, streamId, headerRangeArr, bodyRangeArrItem) {
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    $.ajax({
        url: accountObj.apiUrl + '/objects/' + objectId,
        contentType: "application/json; charset=utf-8",
        method: 'get',
        async: false,
        success: function (res) {
            if (res.success) {
                var object = checkObjectForUpload(res.resource.properties, headerRangeArr, bodyRangeArrItem);
                if (object) {
                    doEditObject(objectId, object);
                } else {
                    sessionStorage.setItem('uploadNum', Number(sessionStorage.getItem('uploadNum')) - 1);
                }
                if (Number(sessionStorage.getItem('uploadNum')) === 0) {
                    BroadcastMessage(streamId);
                    hideStreamsLoading(streamId);
                    showMsg('success', 'Objects upload successfly','#myAccountsMsg');
                }
            }
        },
        error: function (e) {
            sessionStorage.setItem('uploadNum', Number(sessionStorage.getItem('uploadNum')) - 1);
            if (Number(sessionStorage.getItem('uploadNum')) === 0) {
                showMsg('danger', 'Objects upload falied', '#myAccountsMsg');
            }
            //hideStreamsLoading(streamId);
            console.log(e);
        }
    });
}

function doEditObject(objectId, object) {
    var num = Number(sessionStorage.getItem('uploadNum'));
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    var query = {
        properties: object
    }
    $.ajax({
        url: accountObj.apiUrl + '/objects/' + objectId,
        contentType: "application/json; charset=utf-8",
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "Cache-Control": "no-cache"
        },
        async: false,
        processData: false,
        data: JSON.stringify(query),
        method: 'PUT',
        success: function (res) {
            if (res.success) {
                sessionStorage.setItem('uploadNum', Number(sessionStorage.getItem('uploadNum')) - 1);
                showMsg('success', 'Objects ' + objectId + ' upload successfly', '#myAccountsMsg');
            }
            if (Number(sessionStorage.getItem('uploadNum')) === 0) {
                showMsg('success', 'All objects upload successfly', '#myAccountsMsg');
            }
        },
        error: function (e) {
            console.log(e);
            showMsg('danger', 'Objects upload failed', '#myAccountsMsg');
        }
    });
}

function checkObjectForUpload(res_Object, headerRangeArr, rangeArr) {
    var object = {};
    for (let o = 0; o < headerRangeArr.length; o++) {
        object[headerRangeArr[o]] = rangeArr[o];
    }
    // check the columns
    var resKeys = Object.keys(res_Object);
    if (resKeys.length > headerRangeArr.length) {
        return object;
    } else if (resKeys.length < headerRangeArr.length) {
        return object;
    } else {
        for (var i = 0; i < resKeys.length; i++) {
            if (resKeys[i] != headerRangeArr[i]) {
                return object;
            }
        }
    }
    
    // check the data

    for (var key in res_Object) {
        if (res_Object[key] != object[key]) {
            return object;
        }
    }
    return false;
}

// table uploaded auto upload streams
function uploadStreams(resources,streamObj) {
    //streamObj({ streamId: streamId, objects: objects, headerRangeArr: headerRangeArr, bodyRangeArr: bodyRangeArr })
    var uploadsObjects = [];
    if (streamObj && streamObj.objects) {
        uploadsObjects = JSON.parse(JSON.stringify(streamObj.objects));
    }
    if (resources) {
        for (var i = 0; i < resources.length; i++) {
            uploadsObjects.push(resources[i]);
        }
    }

    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    var tableName = sessionStorage.getItem('tableName');
    var query = {
        objects: uploadsObjects,
        layers: [{
            name: tableName,
            orderIndex: 0,
            startIndex: 0,
            objectCount: uploadsObjects.length,
            topology: "0;0-" + uploadsObjects.length,
            guid: getGUID()
        }]
    }
    $.ajax({
        url: accountObj.apiUrl + '/streams/' + streamObj.streamId,
        contentType: "application/json; charset=utf-8",
        method: 'PUT',
        async: false,
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "Cache-Control": "no-cache"
        },
        processData: false,
        data: JSON.stringify(query),
        success: function (res) {
            if (res.success) {
                setTimeout(function () {
                    showMsg('success', 'Streams has been modify successfully', '#myAccountsMsg');
                },3000)
            } else {
                showMsg('danger', 'Unknown error modified streams failed, please try again later', '#myAccountsMsg');
            }
        },
        error: function (e) {
            setTimeout(function () {
                showMsg('danger', 'Unknown error modified streams failed, please try again later', '#myAccountsMsg');
            }, 3000)
            console.log(e);
        }
    });
}

function deleteObject(objectId) {
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    var url = accountObj.apiUrl + '/objects/' + objectId;
    $.ajax({
        url: url,
        contentType: "application/json; charset=utf-8",
        headers: {
            "Content-Type": "application/json",
            "Authorization": accountObj.apiToken,
            "Cache-Control": "no-cache"
        },
        async: false,
        processData: false,
        method: 'DELETE',
        success: function (res) {
            if (res.success === true) {
                showMsg('success', 'Object deleted successfully', '#myAccountsMsg');
            }
        },
        error: function (e) {
            showMsg('danger', 'Object deleted failed', '#myAccountsMsg');
            console.log(e);
        }

    });
}
    

function checkObjects(streamId,objects, headerRangeArr, bodyRangeArr) {
    var queryObjects = JSON.parse(JSON.stringify(objects));

    // Check for new add objects
    if (objects.length > bodyRangeArr.length) {
        queryObjects.splice(bodyRangeArr.length, objects.length - bodyRangeArr.length);
        for (var i = bodyRangeArr.length; i < objects.length; i++) {
            deleteObject(objects[i]._id);
            delete queryObjects[i];
        }
        uploadStreams('', { streamId: streamId, objects: queryObjects });
        uploadObjects(streamId, queryObjects, headerRangeArr, bodyRangeArr);
    } else if (bodyRangeArr.length > objects.length) {
        var addObjects = [];
        for (var j = objects.length; j < bodyRangeArr.length; j++) {
            var row = {};
            for (var h = 0; h < headerRangeArr.length; h++) {
                row[headerRangeArr[h]] = bodyRangeArr[j][h];
            }
            addObjects.push({
                type: "Object",
                properties: row,
                applicationId: 'object' + j + Math.random()
            });
        }
        createdObjects(addObjects, '', { streamId: streamId, objects: objects, headerRangeArr: headerRangeArr, bodyRangeArr: bodyRangeArr });
        uploadObjects(streamId,queryObjects, headerRangeArr, bodyRangeArr);
    } else {
        uploadObjects(streamId, objects, headerRangeArr, bodyRangeArr);
    }
    
}

function checkBindStreamsObjAccount(streamId){
    if (!streamId) {
        return;
    }
    var bindStreamsObj = JSON.parse(getLocalStorage('bindStreamsObj'));
    if (bindStreamsObj[streamId]) {
        var streamsList = JSON.parse(getLocalStorage('streamsList'));
        if (streamsList[streamId].accountObj) {
            sessionStorage.setItem('accountObj', JSON.stringify(streamsList[streamId].accountObj));
        }
    }
}

// When data in the table is changed, this event is triggered.
function onBindingDataChanged(eventArgs) {

    Excel.run(function (ctx) {
        var binding = ctx.workbook.bindings.getItem(eventArgs.binding.id);
        var table = binding.getTable();
        var headerRange = table.getHeaderRowRange().load("values");
        var bodyRange = table.getDataBodyRange().load("values");
        table.load('name');
        binding.load('id');
        return ctx.sync().then(function () {
            showStreamsLoading(binding.id);
            setTimeout(function () {
                getStreamsDetail(binding.id, headerRange.values, bodyRange.values);
            }, 500);
            
        });
    }).catch(ExcelError);
} 

function WebSocketTest(url, streamId) {
    if (!url) {
        return;
    }
    console.log(streamId);
    var msgJson = {
        eventName: "broadcast",
        resourceType: "stream",
        resourceId: streamId,
        args: {
            event: 'update-global'
        },
    };
    var joinJson = {
        eventName: "join",
        resourceType: "stream",
        resourceId: streamId
    };

    var socketStatus = false;
    var myurl = url ? url.replace(/https|http/, 'wss'):'';

    if (!ws[url]) {
        createWebSocket();
    } else if (ws[url].readyState === 1) {
        sendMsg(msgJson);
    } else {
        reconnect();
    }

    // 发送消息
    function sendMsg(myMsg) {
        ws[url].send(JSON.stringify(myMsg) || 'test');
        console.log("数据发送中...");
        console.log(myMsg);
        if (myMsg.eventName != 'alive') {
            setTimeout(function () {
                showMsg('success', 'Upload stream msg transmission ...', '#myAccountsMsg', true);
            }, 500)
        }
    }

    // 重连
    function reconnect() {
        setTimeout(function () {     //没连接上会一直重连，设置延迟避免请求过多
            createWebSocket();
        }, 5000);
    }

    // 实例websocket
    function createWebSocket() {
        if ('WebSocket' in window) {
            console.log("您的浏览器支持 WebSocket!");
            ws[url] = new WebSocket(myurl);

            ws[url].onopen = function () {
                ws[url].sendNumber = 1;
                // Web Socket 已连接上
                sendMsg(joinJson);
            };

            ws[url].onmessage = function (evt) {
                if (evt.data && ws[url].sendNumber == 1) {
                    ws[url].sendNumber++;
                    sendMsg(msgJson);
                    socketStatus = true;
                } else if (evt.data === "ping") {
                    if (socketStatus) {
                        setTimeout(function () {
                            showMsg('success', 'Successful sending of Websocket for update stream', '#myAccountsMsg');
                        }, 800)
                    };
                    socketStatus = false;
                    sendMsg({ eventName: 'alive'});
                    
                }
                console.log("数据已接收...");
            };

            ws[url].onerror = function (evt) {
                console.log(evt);
                console.log("连接错误...");
                reconnect();
            }

            ws[url].onclose = function () {
                // 关闭 websocket
                console.log("连接已关闭...");
            };
        } else {
            console.log("您的浏览器不支持 WebSocket!");
        }
    }
}

function BroadcastMessage(streamId) {
    var accountObj = JSON.parse(getSessionStorage('accountObj'));
    var streamsList = JSON.parse(getLocalStorage('streamsList'));
    var client_id = streamsList[streamId].clientId; 
    //var webUrl = apiAddress + '?client_id=' + client_id +'&access_token=' + accountObj.apiToken;
    var webUrl = accountObj.apiUrl + '?client_id=' + client_id + '&access_token=' + accountObj.apiToken;
    
    WebSocketTest(webUrl, streamId);
}
