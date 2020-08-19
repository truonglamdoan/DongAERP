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
    //$($('a[href="http://www.telerik.com/purchase/aspnet-mvc"]').parent()).css('display', 'inline');
    //$($('a[href="http://www.telerik.com/purchase/aspnet-mvc"]')).css('display', 'none');
    //document.body.innerHTML = document.body.innerHTML.replace("You're using a trial version of Telerik UI for ASP.NET MVC by Progress. Purchase the commercial version now from", "");

    //let el = $('#page-top');
    //el.html(el.html().replace("You're using a trial version of Telerik UI for ASP.NET MVC by Progress. Purchase the commercial version now from", ""));
    //$($('a[href="http://www.telerik.com/purchase/aspnet-mvc"]').parent()).css('display', 'inline');
    //$($('a[href="http://www.telerik.com/purchase/aspnet-mvc"]')).css('display', 'none');


    // Get value từ localStogra
    valueReportType = localStorage.reportTypeLS;

    //DayReport.GetAdditionalData();
    Report.EventDayReport();
    Report.EventMonthReport();
    Report.EventYearReport();
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
    this.EventDayReport = function () {

        // For day report
        let urlGrid = "/Admin/Report/SearchReportDay?fromDay=";
        let urlChart = "/Admin/Report/SearchLineChartReport?fromDay=";
        
        $('ul.search-item .search-grid').click(function () {

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

                let grid = $("#gridReportDay").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                let chart = $("#chartReportDay").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromDateConvert + "&toDay=" + toDateConvert+ "&reportTypeID=" + valueReportType;
                chart.dataSource.read();
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép tìm kiếm trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });

        // PrintExcel for report day
        $('ul.search-item .print-excel').click(function () {

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
                
                window.location = "/ReportLayout/CreateExcel/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType;
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
    this.EventMonthReport = function () {

        // For month report
        let urlGrid = "/Admin/Report/SearchReportMonth?fromDate=";
        let urlChart = "/Admin/Report/SearchLineChartReportMonth?fromDate=";

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

                let grid = $("#gridReportMonth").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                let chart = $("#chartReportMonth").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType;
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
            let fromMonth = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            let toMonth = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số tháng chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24 * 30);

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportLayout/CreateExcel/?fromDate=" + fromMonth + "&toDate=" + toMonth + "&typeID=2" + "&reportTypeID=" + valueReportType;
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
    this.EventYearReport = function () {

        // For month report
        let urlGrid = "/Admin/Report/SearchReportYear?fromYear=";
        let urlChart = "/Admin/Report/SearchLineChartReportYear?fromYear=";

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
                let grid = $("#gridReportYear").data("kendoGrid");
                grid.dataSource.transport.options.read.url = urlGrid + fromYearConvert + "&toYear=" + toYearConvert + "&reportTypeID=" + valueReportType;
                grid.dataSource.read();

                let chart = $("#chartReportYear").data("kendoChart");
                chart.dataSource.transport.options.read.url = urlChart + fromYearConvert + "&toYear=" + toYearConvert + "&reportTypeID=" + valueReportType;
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

                window.location = "/ReportLayout/CreateExcel/?fromDate=" + fromYearConvert + "&toDate=" + toYearConvert + "&typeID=3" + "&reportTypeID=" + valueReportType;
            }
        });
    }

    this.GradationCompartion = function () {

        $('ul.search-item .search-grid-gradation').click(function () {

            //// For month report
            //let urlChart = "/Admin/Report/SearchChartGradationCompare?Gradation=";

            let gradationDicID = $("#gradation").data("kendoComboBox").value();

            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            // Grid doanh số hiển thị
            let grid = $("#gridGradationCompare").data("kendoGrid");
            let urlGrid = "/Admin/Report/SearchGradationCompare?Gradation=";
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID + "&toYear=" + toYear + "&reportTypeID=" + valueReportType;
            grid.dataSource.read();

            // Grid tỉ trọng hiển thị
            let gridGradation = $("#gridGradationComparePercent").data("kendoGrid");
            let urlGridGradation = "/Admin/Report/SearchGradationComparePercentGrid?Gradation=";
            gridGradation.dataSource.transport.options.read.url = urlGridGradation + gradationDicID + "&toYear=" + toYear + "&reportTypeID=" + valueReportType;
            gridGradation.dataSource.read();

            // Biểu đồ cột doanh số
            let chartGradation = $("#chartGradationCompare").data("kendoChart");
            let urlChartDS = "/Admin/Report/SearchDataChartCompare?Gradation=";
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartGradation.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế các tháng hiện tại
            let chartPieProportion = $("#chartGradationPercentYear").data("kendoChart");
            let urlChartPieProportion = "/Admin/Report/SearchPieCompareProportionYear?Gradation=";
            chartPieProportion.dataSource.transport.options.read.url = urlChartPieProportion + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartPieProportion.options.title.text = "Lũy kế " + $("#gradation").data("kendoComboBox").text() + " " + toYear;
            chartPieProportion.dataSource.read();

            // Biểu đồ tròn tỉ trọng theo lũy kế cùng kì năm ngoái
            let chartPieProportionLastYear = $("#chartGradationPercentLastYear").data("kendoChart");
            let urlChartPieProportionLastYear = "/Admin/Report/SearchPieCompareProportionLastYear?Gradation=";
            let lastyear = toYear - 1;
            chartPieProportionLastYear.dataSource.transport.options.read.url = urlChartPieProportionLastYear + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartPieProportionLastYear.options.title.text = "Lũy kế " + $("#gradation").data("kendoComboBox").text() + " " + lastyear;
            chartPieProportionLastYear.dataSource.read();
        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation').click(function () {

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportExcelCompare/CreateExcelGradationCompare/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType;
        });
    }

    this.CompareMonth = function () {

        $('ul.search-item .search-grid-gradation-month').click(function () {

            // For month report

            //let urlChart = "/Admin/Report/SearchDataChartCompareMonth?month=";

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            // grid doanh số
            let urlGridDS = "/Admin/Report/SearchDataGridCompareMonth?Month=";
            let grid = $("#gridGradationCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            grid.dataSource.read();

            // grid tỉ trọng
            let urlGridTT = "/Admin/Report/SearchGridMonthCompareProportion?Month=";
            let gridProportion = $("#gridGradationCompareProportion").data("kendoGrid");
            gridProportion.dataSource.transport.options.read.url = urlGridTT + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            gridProportion.dataSource.read();

            // Biểu đồ cột doanh số
            let chart = $("#chartCompareMonth").data("kendoChart");
            let urlChartDS = "/Admin/Report/SearchDataChartCompareMonth?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chart.dataSource.read();

            // Biểu đồ tròn doanh số theo tháng hiện tại
            let chartPieMonth = $("#chartGradationCompareMonth").data("kendoChart");
            let urlChartPieMonth = "/Admin/Report/SearchPieMonthCompareProportion?month=";
            chartPieMonth.dataSource.transport.options.read.url = urlChartPieMonth + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chartPieMonth.options.title.text = "Tháng " + month + "/" + year;
            chartPieMonth.dataSource.read();

            // Biểu đồ tròn doanh số theo tháng trước
            let chartPieLastMonth = $("#chartGradationCompareLastMonth").data("kendoChart");
            let urlChartPieLastMonth = "/Admin/Report/SearchPieLastMonthCompareProportion?month=";
            chartPieLastMonth.dataSource.transport.options.read.url = urlChartPieLastMonth + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            let lasmonth = month - 1;
            chartPieLastMonth.options.title.text = "Tháng " + lasmonth + "/" + year;
            chartPieLastMonth.dataSource.read();

            // Biểu đồ tròn doanh số cùng kì năm ngoái
            let chartPieMonthLastYear = $("#chartGradationCompareMonthLastYear").data("kendoChart");
            let urlChartPieMonthLastYear = "/Admin/Report/SearchPieMonthLastYearCompareProportion?month=";
            chartPieMonthLastYear.dataSource.transport.options.read.url = urlChartPieMonthLastYear + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chartPieMonthLastYear.options.title.text = "Tháng " + month + "/" + (year - 1);
            chartPieMonthLastYear.dataSource.read();
        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation-month').click(function () {
            
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;
            
            window.location = "/ReportExcelCompare/CreateExcelGradationCompareLastYear/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
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