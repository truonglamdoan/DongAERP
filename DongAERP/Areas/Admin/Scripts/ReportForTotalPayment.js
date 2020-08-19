// #################################################################
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
    // remove telerik của kendo
    $('a[href="http://www.telerik.com/purchase/aspnet-mvc"]').parent().css('display', 'none');

    // Get value từ localStogra
    valueReportType = localStorage.reportTypeLS;

    //DayReport.GetAdditionalData();
    Report.ReportTotalPaymentForDay();
    Report.ReportTotalPaymentForMonth();
    Report.ReportTotalPaymentForYear();
    Report.ReportGradationComapre();
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
    this.ReportTotalPaymentForDay = function () {

        // For day report
        let urlGrid = "/Admin/ReportForTotalPayment/SearchReportTotalPaymentForDay?fromDay=";
        let urlChart = "/Admin/ReportForTotalPayment/SearchLineChartTotalPaymentForReport?fromDay=";

        $('ul.search-item .search-grid-day').click(function () {

            let fromdate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
            let todate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");
            // Tính số ngày chênh lệch
            let difference_In_Time = todate.getTime() - fromdate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

            // loadding
            displayLoading(document.body);

            // Chỉ cho phép lấy thời gian trong khoản 30 ngày
            if ((fromdate.getMonth() == todate.getMonth() && fromdate.getFullYear() == todate.getFullYear())
                // Trường hợp khác tháng cùng năm
                || (todate.getMonth() == fromdate.getMonth() + 1 && fromdate.getFullYear() == todate.getFullYear() && difference_In_Days < 30)
                // Trường hợp khác năm khác tháng
                || fromdate.getMonth() == 11 && fromdate.getFullYear() + 1 == todate.getFullYear() && difference_In_Days < 30) {


                let grid = $("#gridReportTotalPaymentForDay").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;                
                grid.dataSource.read();

                let chart = $("#lineChartReportTotalPaymentForDay").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                chart.options.title.text = "Doanh số chi trả theo tổng doanh số chi trả";
                chart.dataSource.read();
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép tìm kiếm trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });

        // PrintExcel for report day
        $('ul.search-item .print-excel-day').click(function () {

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

                window.location = "/ReportExcelForTotalPayment/CreateExcelForDayMonthYear/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType;
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
    this.ReportTotalPaymentForMonth = function () {

        // For month report
        let urlGrid = "/Admin/ReportForTotalPayment/SearchReportTotalPaymentForMonth?fromDate=";
        let urlChart = "/Admin/ReportForTotalPayment/SearchLineChartReportTotalPaymentForMonth?fromDate=";

        $('ul.search-item .search-grid-month').click(function () {

            let fromDate = $('#FromMonth').data('kendoDatePicker').value();
            let fromMonthConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            let toMonthConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số tháng chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24 * 30);

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                // loadding
                displayLoading(document.body);

                let grid = $("#gridReportTotalPaymentForMonth").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                let chart = $("#lineChartReportTotalPaymentForMonth").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                //chart.options.title.text = "Tổng doanh số chi trả";
                chart.dataSource.read();

            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép tìm kiếm trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        });

        // PrintExcel for report month
        $('ul.search-item .print-excel-month').click(function () {
            let fromDate = $('#FromMonth').data('kendoDatePicker').value();
            let fromMonthConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            let toMonthConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số tháng chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24 * 30);

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportExcelForTotalPayment/CreateExcelForDayMonthYear/?fromDate=" + fromMonthConvert + "&toDate=" + toMonthConvert + "&typeID=2" + "&reportTypeID=" + valueReportType;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo year
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportTotalPaymentForYear = function () {

        // For year report
        let urlGrid = "/Admin/ReportForTotalPayment/SearchReportTotalPaymentForYear?fromDate=";
        let urlChart = "/Admin/ReportForTotalPayment/SearchLineChartTotalPaymentReportForYear?fromDate=";

        $('ul.search-item .search-grid-year').click(function () {
            
            let fromYear = $('#FromYear').data('kendoDatePicker').value();
            let fromYearConvert = kendo.toString(fromYear, "MM/dd/yyyy");
            let toYear = $('#ToYear').data('kendoDatePicker').value();
            let toYearConvert = kendo.toString(toYear, "MM/dd/yyyy");

            // Chỉ show dc 5 năm
            if (toYear.getFullYear() - fromYear.getFullYear() > 4) {

                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép tìm kiếm trong 5 năm trở lại"
                }).data("kendoAlert").open();
            } else {
                let grid = $("#gridReportTotalPaymentForYear").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;

                //// thay đổi giá trị của cột
                //let yearStart = fromYear.getFullYear();
                //columns = grid.columns;
                //let count = 0;
                //for (let i = 1; i < columns.length; i++) {
                //    // fied của cột
                //    let fied = grid.columns[i].field;
                //    titleName = yearStart + count;
                //    $("#gridReportTotalPaymentForYear thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                //    count++;
                //}

                // loadding
                displayLoading(document.body);

                grid.dataSource.read();

                // Search cho line char
                let chart = $("#lineChartReportTotalPaymentForYear").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;
                //chart.options.title.text = "Tổng doanh số chi trả";
                chart.dataSource.read();
            }
        });

        // PrintExcel for report year
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

                window.location = "/ReportExcelForTotalPayment/CreateExcelForDayMonthYear/?fromDate=" + fromYearConvert + "&toDate=" + toYearConvert + "&typeID=3" + "&reportTypeID=" + valueReportType;
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo year
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportGradationComapre = function () {

        $('ul.search-item .search-grid-gradation').click(function () {
            
            // loadding
            displayLoading(document.body);

            // For year report
            let urlGrid = "/Admin/ReportForTotalPayment/SearchReportGradationCompare?gradation=";
            let gradationDicID = $("#gradation").data("kendoComboBox");
            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            let grid = $("#gridGradationCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID.value() + "&year=" + toYear + "&reportTypeID=" + valueReportType;  
            grid.dataSource.read();

            let chart = $("#chartGradationCompare").data("kendoChart");
            let urlChart = "/Admin/ReportForTotalPayment/SearchColumnChartMaketReportForGradation?gradation=";

            chart.dataSource.transport.options.read.url = urlChart + gradationDicID.value() + "&year=" + toYear + "&reportTypeID=" + valueReportType;  
            chart.options.title.text = kendo.format("Tổng doanh số chi trả \n giai đoạn: {0} {1}", gradationDicID.text(), toYear);
            chart.dataSource.read();
        });

        // Create report Excel for theo giai đoàn
        $('ul.search-item .print-excel-gradation').click(function () {

            let gradation = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportExcelForTotalPayment/CreateExcelForGradationCompare/?gradation=" + gradation + "&year=" + year + "&reportTypeID=" + valueReportType;  
        });
    }


    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo year
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.CompareMonth = function () {

        $('ul.search-item .search-grid-gradation-month').click(function () {

            // For month report

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;
            // loadding
            displayLoading(document.body);
            
            let urlGridDS = "/Admin/ReportForTotalPayment/SearchReportCompareForMonth?month=";
            let grid = $("#gridGradationCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;

            grid.dataSource.read();
            
            // Biểu đồ cột doanh số
            let chart = $("#chartCompareMonth").data("kendoChart");
            let urlChartDS = "/Admin/ReportForTotalPayment/SearchColumnsChartCompareForMonth?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chart.options.title.text = kendo.format("Tổng doanh số chi trả tháng {0}/{1} \n so với tháng trước và cùng kì năm trước", month, year);
            chart.dataSource.read();
        });

        // PrintExcel for report month
        $('ul.search-item .print-excel-gradation-month').click(function () {

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            // Trường hợp type= 2 để get đúng dữ liệu trong Db
            window.location = "/ReportExcelForTotalPayment/CreateExcelGradationCompareLastYear/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
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