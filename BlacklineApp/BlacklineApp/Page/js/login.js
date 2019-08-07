$(function() {
    loadingFrm()
})
function logon() {
    password = $('#password').val()
    username = $('#username').val()
    if (password == '' || username == '') {
        showMessage('Please input username and password')
    }
    else {
        //showLoading()
        checkUserInfo(username, password)
    }
}

function showChromeTool() {
    if ($(" input[ name='name' ] ").val() == 999991) {
        showDevToolsFrm()
    }
    if ($(" input[ name='name' ] ").val() == 888881) {
        closeDevToolsFrm()
    }
    return
}

function redirectiUrl(url) {
    window.location.href = url
}

function hideLoading() {
    $('#login_wait').modal('hide');
}

function showLoading() {
    $('#login_wait').modal('show');
}

function showMessage(msg, title = 'Info', width = 400) {
    msg = decodeURIComponent(msg)
    var d = dialog({
        width: width,
        //height:80,
        title: title,
        content: msg,
        ok: function () { },
        //statusbar: '<label><input type="checkbox">不再提醒</label>'
    });
    d.show();
    //alert(msg)
}