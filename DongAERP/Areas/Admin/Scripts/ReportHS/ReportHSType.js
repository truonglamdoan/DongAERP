﻿// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	18/12/2019		Trường Lãm		Create New
// ##################################################################

// Nếu là reportDay : type = 1, report month : type = 2, report year: type = 3
var typeID = null;
var valueReportType = null;

$(document).ready(function () {

    // Get value từ localStogra
    valueReportType = localStorage.reportTypeLS;

    //DayReport.GetAdditionalData();
    Report.ReportDay();
    Report.ReportMonth();
    Report.ReportYear();
    // Search báo cáo so sánh theo giai đoạn
    Report.GradationCompartion();
    // Search báo cáo so sánh theo tháng
    Report.CompareMonth();

    // sự kiện click event chọn Type
    // Change radio
    $("input[name='reportDate']").change(function () {

        // Get value từ localStogra
        valueReportType = localStorage.reportTypeLS;
    });
});

var Report = new function () {

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportDay = function () {
        
        // PrintExcel for report day
        $('ul.search-item .print-excel-day').click(function () {

            // loadding
            displayLoading(document.body);

            let fromdate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
            let todate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");

            // Tính số ngày chênh lệch
            let difference_In_Time = todate.getTime() - fromdate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

            // Chỉ cho phép lấy thời gian trong khoản 30 ngày
            if ((fromdate.getMonth() == todate.getMonth() && fromdate.getFullYear() == todate.getFullYear())
                // Trường hợp khác tháng cùng năm
                || (todate.getMonth() == fromdate.getMonth() + 1 && fromdate.getFullYear() == todate.getFullYear() && difference_In_Days < 30)
                // Trường hợp khác năm khác tháng
                || fromdate.getMonth() == 11 && fromdate.getFullYear() + 1 == todate.getFullYear() && difference_In_Days < 30) {
                
                window.location = "/ReportHSTypeExcel/CreateExcel/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo tháng
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportMonth = function () {
        
        // PrintExcel for report month
        $('ul.search-item .print-excel-month').click(function () {
            let fromDate = $('#FromMonth').data('kendoDatePicker').value();
            let fromMonth = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            let toMonth = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số tháng chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24 * 30);

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportHSTypeExcel/CreateExcel/?fromDate=" + fromMonth + "&toDate=" + toMonth + "&typeID=2" + "&reportTypeID=" + valueReportType;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo năm
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportYear = function () {
        
        $('ul.search-item .print-excel-year').click(function () {
            let fromYear = $('#FromYear').data('kendoDatePicker').value();
            let fromYearConvert = kendo.toString(fromYear, "MM/dd/yyyy");
            let toYear = $('#ToYear').data('kendoDatePicker').value();
            let toYearConvert = kendo.toString(toYear, "MM/dd/yyyy");

            // Chỉ show dc 5 năm
            if (toYear.getFullYear() - fromYear.getFullYear() > 4) {

                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 5 năm trở lại"
                }).data("kendoAlert").open();
            } else {

                window.location = "/ReportHSTypeExcel/CreateExcel/?fromDate=" + fromYearConvert + "&toDate=" + toYearConvert + "&typeID=3" + "&reportTypeID=" + valueReportType;
            }
        });
    }

    this.GradationCompartion = function () {
        
        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation').click(function () {

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportHSTypeExcel/CreateExcelGradationCompare/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType;
        });
    }

    this.CompareMonth = function () {
        
        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-compare-month').click(function () {
            
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;
            
            window.location = "/ReportHSTypeExcel/CreateExcelCompareMonth/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
        });
    }
}

// Loading
function displayLoading(target) {

    var element = $(target);
    kendo.ui.progress(element, true);
    setTimeout(function () {
        kendo.ui.progress(element, false);
    }, 2000);
}