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
    //DayReport.GetAdditionalData();
    Report.ReportTotalMoneyTypeForDay();
    Report.ReportTotalMoneyTypeForMonth();
    Report.ReportTotalMoneyTypeForYear();
    Report.ReportGradationComapre();
    Report.CompareMonth();

    // Get value từ localStogra
    valueReportType = localStorage.reportTypeLS;
    
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
    this.ReportTotalMoneyTypeForDay = function () {

        // For day report
        let urlGrid = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForDay?fromDay=";
        let urlGridConvert = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForDayConvert?fromDay=";
        let urlChart = "/Admin/ReportTotalMoneyType/SearchLineChartTotalMoneyTypeForReport?fromDay=";

        $('ul.search-item .search-grid-day').click(function () {

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
                
                // loadding
                displayLoading(document.body);

                let grid = $("#gridReportTotalMoneyTypeForDay").data("kendoGrid");
                
                grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                let gridConvert = $("#gridReportTotalMoneyTypeForDayConvert").data("kendoGrid");
                gridConvert.dataSource.transport.options.read.url = urlGridConvert + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                gridConvert.dataSource.read();

                // Line chart cho report day
                let chart = $("#lineChartReportTotalMoneyTypeForDay").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                chart.options.title.text = "Doanh số chi trả theo từng loại tiền";
                chart.dataSource.read();

                // Column chart cho report Day
                let chartColumn = $("#ColumnchartForDay").data("kendoChart");
                let urlChartColumn = "/Admin/ReportTotalMoneyType/SearchColumnChartTotalMoneyTypeForReport?fromDay=";
                chartColumn.dataSource.transport.options.read.url = urlChartColumn + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                chartColumn.options.title.text = "Doanh số chi trả theo từng loại tiền";
                chartColumn.dataSource.read();
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

                window.location = "/ReportExcelForTotalMoneyType/CreateExcelForDayMonthYear/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType;
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
    this.ReportTotalMoneyTypeForMonth = function () {


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

                // search for month theo nguyên tệ
                let grid = $("#gridReportTotalMoneyTypeForMonh").data("kendoGrid");
                // For month report
                let urlGrid = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForMonth?fromDate=";
                grid.dataSource.transport.options.read.url = urlGrid + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                // search for month theo quy đổi USD
                let gridConvert = $("#gridReportTotalMoneyTypeForMonhConvert").data("kendoGrid");
                // For month report
                let urlGridConvert = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForMonthConvert?fromDate=";
                gridConvert.dataSource.transport.options.read.url = urlGridConvert + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                gridConvert.dataSource.read();

                // Biểu đồ đường
                let urlChart = "/Admin/ReportTotalMoneyType/SearchLineChartReportTotalMoneyTypeForMonh?fromDate=";
                let chart = $("#lineChartReportTotalMoneyTypeForMonh").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                //chart.options.title.text = "Doanh số chi trả theo từng loại tiền";
                chart.dataSource.read();

                // Biểu đồ cột chồng
                let urlColumnChart = "/Admin/ReportTotalMoneyType/SearchColumnChartReportTotalMoneyTypeForMonh?fromDate=";
                let columnChart = $("#ColumnchartForMonh").data("kendoChart");
                columnChart.dataSource.transport.options.read.url = urlColumnChart + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                //columnChart.options.title.text = "Doanh số chi trả theo từng loại tiền";
                columnChart.dataSource.read();

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
            fromMonth = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            toMonth = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số tháng chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24 * 30);

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportExcelForTotalMoneyType/CreateExcelForDayMonthYear/?fromDate=" + fromMonth + "&toDate=" + toMonth + "&typeID=2" + "&reportTypeID=" + valueReportType;
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
    this.ReportTotalMoneyTypeForYear = function () {

        // For year report

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

                // loadding
                displayLoading(document.body);

                // bảng nguyển tệ
                let urlGrid = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForYear?fromDate=";
                let grid = $("#gridReportTotalMoneyTypeForYear").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                // Bảng quy đổi USD
                let urlGridConvert = "/Admin/ReportTotalMoneyType/SearchReportTotalMoneyTypeForYearConvert?fromDate=";
                let gridConvert = $("#gridReportTotalMoneyTypeForYearConvert").data("kendoGrid");
                gridConvert.dataSource.transport.options.read.url = urlGridConvert + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;
                gridConvert.dataSource.read();

                // Search dữ liệu cho line chart
                let urlChart = "/Admin/ReportTotalMoneyType/SearchLineChartTotalMoneyTypeReportForYear?fromDate=";
                let chart = $("#lineChartReportTotalMoneyTypeForYear").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;
                //chart.options.title.text = "Doanh số chi trả theo từng loại tiền";
                chart.dataSource.read();

                // Search dữ liệu cho column chart
                let urlColumnChart = "/Admin/ReportTotalMoneyType/SearchColumnChartGradationCompareForYear?fromDate=";
                let columnChart = $("#ColumnchartForYear").data("kendoChart");
                columnChart.dataSource.transport.options.read.url = urlColumnChart + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType;
                //columnChart.options.title.text = "Doanh số chi trả theo từng loại tiền";
                columnChart.dataSource.read();
            }

        });

        // PrintExcel for report year
        $('ul.search-item .print-excel-year').click(function () {
            let fromYear = $('#FromYear').data('kendoDatePicker').value();
            fromYear = kendo.toString(fromYear, "MM/dd/yyyy");
            let toYear = $('#ToYear').data('kendoDatePicker').value();
            toYear = kendo.toString(toYear, "MM/dd/yyyy");

            window.location = "/ReportExcelForTotalMoneyType/CreateExcelForDayMonthYear/?fromDate=" + fromYear + "&toDate=" + toYear + "&typeID=3" + "&reportTypeID=" + valueReportType;
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo year
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.ReportGradationComapre = function () {

        $('ul.search-item .search-grid-gradation').click(function () {

            // Theo nguyên tệ
            // For year report
            let urlGrid = "/Admin/ReportTotalMoneyType/SearchReportGradationCompare?gradation=";
            let gradationDicID = $("#gradation").data("kendoComboBox");
            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();
            
            // loadding
            displayLoading(document.body);

            let grid = $("#gridGradationCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID.value() + "&toYear=" + toYear + "&reportTypeID=" + valueReportType;

            // Change text column
            let columns = grid.columns;
            let titleName;
            for (let i = 0; i < columns.length; i++) {
                let fied = grid.columns[i].field;
                // Lũy kế năm hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + toYear.toString();
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Lũy kế năm trước
                if (fied == "AccumulateID2") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + (toYear - 1).toString();
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tăng giảm theo %
                if (fied == "CompareToIDPercent") {
                    titleName = "Tăng giảm </br>so với cùng kỳ " + (toYear - 1).toString() + " (%)";
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tăng giảm theo +/-
                if (fied == "CompareToID") {
                    titleName = "Tăng giảm </br>so với cùng kỳ " + (toYear - 1).toString() + " (+/-)";
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            grid.dataSource.read();

            // Quy USD
            // For year report
            let urlGridConvert = "/Admin/ReportTotalMoneyType/SearchReportGradationCompareConvert?gradation=";
            let gridConvert = $("#gridGradationCompareConvert").data("kendoGrid");
            gridConvert.dataSource.transport.options.read.url = urlGridConvert + gradationDicID.value() + "&toYear=" + toYear + "&reportTypeID=" + valueReportType;

            // Change text column
            let columnsConvert = gridConvert.columns;
            for (let i = 0; i < columnsConvert.length; i++) {
                let fied = gridConvert.columns[i].field;
                // Lũy kế năm hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + toYear.toString();
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Lũy kế năm trước
                if (fied == "AccumulateID2") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + (toYear - 1).toString();
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tăng giảm theo %
                if (fied == "CompareToIDPercent") {
                    titleName = "Tăng giảm </br>so với cùng kỳ " + (toYear - 1).toString() + " (%)";
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tăng giảm theo +/-
                if (fied == "CompareToID") {
                    titleName = "Tăng giảm </br>so với cùng kỳ " + (toYear - 1).toString() + " (+/-)";
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            gridConvert.dataSource.read();

            let chart = $("#chartGradationCompare").data("kendoChart");
            let urlChart = "/Admin/ReportTotalMoneyType/SearchColumnChartReportForGradation?gradation=";

            chart.dataSource.transport.options.read.url = urlChart + gradationDicID.value() + "&year=" + toYear;
            chart.options.title.text = "Doanh số theo từng loại tiền (Quy USD) \n" + "Giai đoạn: " + gradationDicID.text();
            chart.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế các tháng hiện tại
            let chartPieGradationForYear = $("#chartGradationPercentForYear").data("kendoChart");
            let urlChartPieGradationForYear = "/Admin/ReportTotalMoneyType/SearchGradationComparePieForYear?gradation=";
            chartPieGradationForYear.dataSource.transport.options.read.url = urlChartPieGradationForYear + gradationDicID.value() + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartPieGradationForYear.options.title.text = "Lũy kế " + $("#gradation").data("kendoComboBox").text() + " " + toYear;
            chartPieGradationForYear.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế cùng kì năm ngoái
            let chartPieGradationForLastYear = $("#chartGradationPercentLastYear").data("kendoChart");
            let urlChartPieGradationForastYear = "/Admin/ReportTotalMoneyType/SearchGradationComparePieForLastYear?gradation=";
            let lastyear = toYear - 1;
            chartPieGradationForLastYear.dataSource.transport.options.read.url = urlChartPieGradationForastYear + gradationDicID.value() + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartPieGradationForLastYear.options.title.text = "Lũy kế " + $("#gradation").data("kendoComboBox").text() + " " + lastyear;
            chartPieGradationForLastYear.dataSource.read();


            // Search cho bảng dữ liệu phần trăm
            let urlGridGradationComparePercent = "/Admin/ReportTotalMoneyType/SearchReportGradationComparePercent?gradation=";
            let gridurlGridGradationComparePercent = $("#gridGradationComparePercent").data("kendoGrid");
            gridurlGridGradationComparePercent.dataSource.transport.options.read.url = urlGridGradationComparePercent + gradationDicID.value() + "&year=" + toYear + "&reportTypeID=" + valueReportType;

            // Change text column
            columns = gridurlGridGradationComparePercent.columns;
            for (let i = 0; i < columns.length; i++) {
                let fied = gridurlGridGradationComparePercent.columns[i].field;
                // Lũy kế năm hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + toYear.toString();
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Lũy kế năm trước
                if (fied == "AccumulateID2") {
                    titleName = "Lũy kế </br>" + gradationDicID.text() + " " + (toYear - 1).toString();
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tăng giảm theo %
                if (fied == "CompareToIDPercent") {
                    titleName = "Tăng giảm </br> so với cùng kỳ " + (toYear - 1).toString() + " (%)";
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            gridurlGridGradationComparePercent.dataSource.read();
        });

        // Create report Excel for theo giai đoàn
        $('ul.search-item .print-excel-gradation').click(function () {

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportExcelForTotalMoneyType/CreateExcelForGradationCompare/?gradationID=" + gradationID + "&year=" + year + "&typeID=1" + "&reportTypeID=" + valueReportType;
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo year
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.CompareMonth = function () {

        $('ul.search-item .search-grid-gradation-month').click(function () {
            
            // loadding
            displayLoading(document.body);

            // For month report
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;
            
            let urlGrid = "/Admin/ReportTotalMoneyType/SearchReportCompareForMonth?month=";
            let grid = $("#gridGradationCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGrid + month + "&year=" + year;

            // Change text column
            columns = grid.columns;
            for (let i = 0; i < columns.length; i++) {
                let fied = grid.columns[i].field;
                // Tháng hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Tháng </br>" + month + "/" + year;
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tháng trước
                if (fied == "AccumulateID2") {
                    titleName = "Tháng </br>" + (month - 1) + "/" + year;
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Cùng kì năm trước
                if (fied == "AccumulateID3") {
                    titleName = "Tháng </br>" + month + "/" + (year - 1);
                    $("#gridGradationCompare thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            grid.dataSource.read();

            let urlGridConvert = "/Admin/ReportTotalMoneyType/SearchReportCompareForMonthConvert?month=";
            let gridConvert = $("#gridGradationCompareConvert").data("kendoGrid");
            gridConvert.dataSource.transport.options.read.url = urlGridConvert + month + "&year=" + year;

            // Change text column
            columns = gridConvert.columns;
            for (let i = 0; i < columns.length; i++) {
                let fied = gridConvert.columns[i].field;
                // Tháng hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Tháng </br>" + month + "/" + year;
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tháng trước
                if (fied == "AccumulateID2") {
                    titleName = "Tháng </br>" + (month - 1) + "/" + year;
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Cùng kì năm trước
                if (fied == "AccumulateID3") {
                    titleName = "Tháng </br>" + month + "/" + (year - 1);
                    $("#gridGradationCompareConvert thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            gridConvert.dataSource.read();

            
            // Biểu đồ cột doanh số
            let chart = $("#chartCompareMonth").data("kendoChart");
            let urlChartDS = "/Admin/ReportTotalMoneyType/SearchColumnsChartCompareForMonth?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chart.options.title.text = kendo.format("Doanh số chi trả theo từng loại tiền trong tháng {0} \n so với tháng trước và cùng kì năm trước", month);
            chart.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế các tháng hiện tại
            let chartPieCompareForMonth = $("#chartGradationCompareMonth").data("kendoChart");
            let urlChartPieCompareForMonth = "/Admin/ReportTotalMoneyType/SearchGradationComparePieForMonth?month=";
            chartPieCompareForMonth.dataSource.transport.options.read.url = urlChartPieCompareForMonth + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chartPieCompareForMonth.options.title.text = kendo.format("Tháng {0}/{1}", month, year);
            chartPieCompareForMonth.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế các tháng trước
            let chartPieCompareForLastMonth = $("#chartGradationCompareLastMonth").data("kendoChart");
            let urlChartPieCompareForLastMonth = "/Admin/ReportTotalMoneyType/SearchGradationComparePieForMonthLastMonth?month=";
            chartPieCompareForLastMonth.dataSource.transport.options.read.url = urlChartPieCompareForLastMonth + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chartPieCompareForLastMonth.options.title.text = kendo.format("Tháng {0}/{1}", month - 1, year);
            chartPieCompareForLastMonth.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế cùng kì năm trước
            let chartPieCompareForMonthLastYear = $("#chartGradationCompareMonthLastYear").data("kendoChart");
            let urlChartPieCompareForMonthLastYear = "/Admin/ReportTotalMoneyType/SearchGradationComparePieForMonthLastYear?month=";
            chartPieCompareForMonthLastYear.dataSource.transport.options.read.url = urlChartPieCompareForMonthLastYear + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chartPieCompareForMonthLastYear.options.title.text = kendo.format("Tháng {0}/{1}", month, year - 1);
            chartPieCompareForMonthLastYear.dataSource.read();

            // Search bảng dữ liệu theo
            let urlGridCompareForMonth = "/Admin/ReportTotalMoneyType/SearchReportCompareForMonthPercent?month=";
            let gridCompareForMonth = $("#gridGradationComparePercent").data("kendoGrid");
            gridCompareForMonth.dataSource.transport.options.read.url = urlGridCompareForMonth + month + "&year=" + year + "&reportTypeID=" + valueReportType;

            // Change text column
            columns = gridCompareForMonth.columns;
            for (let i = 0; i < columns.length; i++) {
                let fied = gridCompareForMonth.columns[i].field;
                // Tháng hiện tại
                if (fied == "AccumulateID1") {
                    titleName = "Tháng </br>" + month + "/" + year;
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Tháng trước
                if (fied == "AccumulateID2") {
                    titleName = "Tháng </br>" + (month - 1) + "/" + year;
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }

                // Cùng kì năm trước
                if (fied == "AccumulateID3") {
                    titleName = "Tháng </br>" + month + "/" + (year - 1);
                    $("#gridGradationComparePercent thead").find("[data-field='" + fied + "'] .k-link").html(titleName);
                }
            }
            gridCompareForMonth.dataSource.read();
        });

        // PrintExcel for report month
        $('ul.search-item .print-excel-gradation-month').click(function () {

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            // Trường hợp type= 2 để get đúng dữ liệu trong Db
            window.location = "/ReportExcelForTotalMoneyType/CreateExcelForCompareYear/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
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