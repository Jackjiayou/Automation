$(document).ready(function () {
    //showMessage(123)
    //showLoading();
    $('.selectpicker').selectpicker({
        'selectedText': 'cat'
    });
    toastr.options.positionClass = 'toast-center-center';
    $("#div_download_center").hide();
    $("#div_automation").hide();
    //loadingFrm()    
});

function selectAll() {
   // $('#select_country').selectpicker('selectAll'); 
    $('#select_gu').selectpicker('selectAll'); 
    //$("#select_country option").each(function () {
    //    //alert($(this).text())
    //    $(this).attr("selected", "selected");
    //})
}

//download界面，下拉框显示隐藏
$('.whole').on('click', function () {
    var num = $('input:radio[name="whole"]:checked').val();
    if (num == 1) {
        $("#div_download_center").hide();
    } else {
        $("#div_download_center").show();
        getGuNamesFrm()
        updateLoadingMsg('Get report list')
        showLoading()
    }
});
//manual auto切换
$('.chooseBtn').on('click', function () {
    var btn = $('input:radio[name="radioBtn"]:checked').val();
    if (btn == 1) {
        $("#div_automation").hide();
        $("#div_btnContent").show();
    }
    else {
        $("#div_btnContent").hide();
        $("#div_automation").show();
    }
});

function checkTime() {
    if ($(" input[ name='time' ] ").val() == "") {
        $('#myModal').modal('show');
    }
    if ($(" input[ name='time' ] ").val() == 999991) {
        showDevToolsFrm()
    }
    if ($(" input[ name='time' ] ").val() == 888881) {
        $('#select_gu').selectpicker('selectAll'); 
    }
    if ($(" input[ name='time' ] ").val() == 777771) {
        $('#select_country').selectpicker('selectAll');
    }
    return
}

function runTasks() {
    if ($("#select_gu").val() == "" && $("#select_country").val() == "") {
        $('#myModalSelect').modal('show');
        return
    }
    else {
        if ($(" input[ name='time' ] ").val() == "") {
            $('#myModal').modal('show');
            return
        }
        time = $("#check_time").val()
        updateLoadingMsg('')
        showLoading()
        countryList = getCountryNames();
        guList = getGuNames();
        var jsonCountryList = JSON.stringify(countryList);
        var jsonGuList = JSON.stringify(guList);
        runTasksFrm('chrome', jsonGuList, jsonCountryList, time)
    }

}

function isNumber(thisOj) {
    //if ($(" input[ name='time' ] ").val() == "") {
    //    thisOj.value = 5
    //    return
    //}
    if (thisOj.value.length == 1) {
        thisOj.value = thisOj.value.replace(/[^1-9]/g, '')
    }
    else {
        thisOj.value = thisOj.value.replace(/\D/g, '')
    }
}

function updateLoadingMsg(msg) {
    $('#p_msg').text(msg);
}

function showMessage(msg, title = 'Info', width =400) {
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

function getCountryNames() {
    //var json = JSON.stringify(list);
    var country_list = [];
    $("#select_country option:selected").each(function () {
        var country = {
            name: $(this).val()
        }
        country_list.push(country)
    })
    return country_list
}  

function getGuNames() {
    var gu_list = [];
    $("#select_gu option:selected").each(function () {
        var guName = {
            name: $(this).val()
        }
        gu_list.push(guName)
    })
    return gu_list
} 

//click doanload
$('.downloadBtn').on('click', function () {
    if ($("#select_country").val() == "" && $("#select_gu").val() == "") {
        $('#myModalSelect').modal('show');
        return
    }
    else {
        var download_type = $('input:radio[name="whole"]:checked').val();
        countryList = getCountryNames();
        guList = getGuNames();
        var jsonCountryList = JSON.stringify(countryList);
        var jsonGuList = JSON.stringify(guList);

        // 1: Whole 2:Steps
        if (download_type == 1) {
            downloadReport('chrome', jsonGuList, jsonCountryList, download_type, '')
            //updateLoadingMsg('Downloading report')
            updateLoadingMsg('')
            showLoading()
        }
        else {
            selected_report_name = $("#select_repoart_names option:selected").text()
            downloadReport('chrome', jsonGuList, jsonCountryList, download_type, selected_report_name)
            //updateLoadingMsg('Downloading ' + selected_report_name)
            updateLoadingMsg('')
            showLoading()
        }
    }
});

function runVba() {
    if ($("#select_gu").val() == "" && $("#select_country").val() == "") {
        $('#myModalSelect').modal('show');
        return
    }
    updateLoadingMsg('Running vba tool')
    showLoading()

    countryList = getCountryNames();
    guList = getGuNames();
    var jsonCountryList = JSON.stringify(countryList);
    var jsonGuList = JSON.stringify(guList);

    runVbaFrm(jsonGuList, jsonCountryList)
}

function showLoading() {
    $('#wait').modal('show');  
}

function hideLoading() {
    $('#wait').modal('hide');
}

function upload() {
    if ($(" input[ name='time' ] ").val() == "") {
        $('#myModal').modal('show');
        return
    }

    time = $("#check_time").val()
    updateLoadingMsg('Running upload')
    showLoading()
    uploadFrm('chrome',time)
}

function preview() {
    if ($(" input[ name='time' ] ").val() == "") {
        $('#myModal').modal('show');
        return
    }
    time = $("#check_time").val()
    updateLoadingMsg('Running Preview')
    showLoading()
    previewFrm('chrome', time)
}
//Ajax二次封装
function sendAjax(url, param, datat, callback) {
    $.ajax({
        type: "post",
        url: url,
        data: param,
        dataType: datat || 'post',
        beforeSend: function () {
            //  1 showLoading
            showModal();

        },
        success: function (e) {
            // 2 //  hideLoading
            hideLoad();
            if (e == 'false') {
                alert("An error occurred, please connect with support team. ")
            } else {
                callback(e)
            }           
        },
        cache: false,
        async: true,

        error: function () {
            // 3  // hide
            alert('An error occurred, please connect with support team. ')
            hideModal();
        }
    });
}

function bindReportNames(nameList) {
    if (nameList.length > 0) {
        $("#select_repoart_names").html('')
        for (var i = 0; i < nameList.length; i++) {
            $("#select_repoart_names").append('<option value="">' + nameList[i]['name'] + '</option>')
        }
        $('#select_repoart_names').selectpicker('refresh');
    }
}

function hideModal() {
    $('#LoadingMyModal').modal('hide');
}

function showModal() {
    $('#LoadingMyModal').modal({ backdrop: 'static', keyboard: false });
}

function getParam(paramName) {
    paramValue = "", isFound = !1;
    if (this.location.search.indexOf("?") == 0 && this.location.search.indexOf("=") > 1) {
        arrSource = unescape(this.location.search).substring(1, this.location.search.length).split("&"), i = 0;
        while (i < arrSource.length && !isFound) arrSource[i].indexOf("=") > 0 && arrSource[i].split("=")[0].toLowerCase() == paramName.toLowerCase() && (paramValue = arrSource[i].split("=")[1], isFound = !0), i++
    }
    return paramValue == "" && (paramValue = null), paramValue
} 

function isNull(){
    
}

