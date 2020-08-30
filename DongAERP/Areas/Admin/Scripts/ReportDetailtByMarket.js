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
    Report.MarketForTotal();
    Report.MarketForTotalForMonth();
    Report.MarketForTotalForYear();
    
    Report.MarketForOne();
    // Search báo cáo so sánh theo giai đoạn
    Report.GradationCompartion();
    // Search báo cáo so sánh theo giai đoạn theo từng đối tác
    Report.GradationCompartionForOne();
    //// Search báo cáo so sánh theo tháng cho tất cả thị trường
    Report.CompareMonthForAll();

    //// Search báo cáo so sánh theo tháng từng thị trường
    Report.CompareMonthForOne();

    // sự kiện click event chọn Type
    // Change radio
    $("input[name='reportDate']").change(function () {

        // Get value từ localStogra
        valueReportType = localStorage.reportTypeLS;
    });
});

var Report = new function () {

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.MarketForTotal = function () {

        //// For day report
        //let urlGrid = "/Admin/ReportDetailtByMarket/SearchMarketForTotal?fromDay=";
        
        //$('ul.search-item .search-grid-forAll').click(function () {

        //    // loadding
        //    displayLoading(document.body);

        //    // Get mã thị trường
        //    let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

        //    let fromdate = $('#FromDay').data('kendoDatePicker').value();
        //    let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
        //    let todate = $('#ToDay').data('kendoDatePicker').value();
        //    let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");

        //    // Tính số ngày chênh lệch
        //    let difference_In_Time = todate.getTime() - fromdate.getTime();
        //    let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

        //    if (marketID != '') {
        //        // Chỉ cho phép lấy thời gian trong khoản 30 ngày
        //        if ((fromdate.getMonth() == todate.getMonth() && fromdate.getFullYear() == todate.getFullYear())
        //            // Trường hợp khác tháng cùng năm
        //            || (todate.getMonth() == fromdate.getMonth() + 1 && fromdate.getFullYear() == todate.getFullYear() && difference_In_Days < 30)
        //            // Trường hợp khác năm khác tháng
        //            || fromdate.getMonth() == 11 && fromdate.getFullYear() + 1 == todate.getFullYear() && difference_In_Days < 30) {

        //            let grid = $("#gridReportDetailtByTotal").data("kendoGrid");
        //            grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
        //            grid.dataSource.read();

        //        } else {
        //            $("<div></div>").kendoAlert({
        //                title: "Cảnh báo!",
        //                content: "Bạn chỉ được phép tìm kiếm trong 30 ngày trở lại"
        //            }).data("kendoAlert").open();
        //        }
        //    }
        //});

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

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
                
                window.location = "/ReportDetailtExcelForMarket/CreateExcelForMarket/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }


    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.MarketForTotalForMonth = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll-ofMonth').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

            let fromDate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
            let toDate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");
            
            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportDetailtExcelForMarket/CreateExcelForMarket/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.MarketForTotalForYear = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll-ofYear').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

            let fromDate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
            let toDate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");

            if (toDate.getFullYear() - fromDate.getFullYear() > 4) {

                window.location = "/ReportDetailtExcelForMarket/CreateExcelForMarket/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.MarketForOne = function () {

        // For day report
        let urlGrid = "/Admin/ReportDetailtByMarket/SearchMarketForOne?fromDay=";

        //$('ul.search-item .search-grid-MarketOne').click(function () {

        //    // loadding
        //    displayLoading(document.body);

        //    // Get mã thị trường
        //    let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

        //    let fromdate = $('#FromDay').data('kendoDatePicker').value();
        //    let fromDateConvert = kendo.toString(fromdate, "MM/dd/yyyy");
        //    let todate = $('#ToDay').data('kendoDatePicker').value();
        //    let toDateConvert = kendo.toString(todate, "MM/dd/yyyy");

        //    // Tính số ngày chênh lệch
        //    let difference_In_Time = todate.getTime() - fromdate.getTime();
        //    let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

        //    if (marketID != '') {
        //        // Chỉ cho phép lấy thời gian trong khoản 30 ngày
        //        if ((fromdate.getMonth() == todate.getMonth() && fromdate.getFullYear() == todate.getFullYear())
        //            // Trường hợp khác tháng cùng năm
        //            || (todate.getMonth() == fromdate.getMonth() + 1 && fromdate.getFullYear() == todate.getFullYear() && difference_In_Days < 30)
        //            // Trường hợp khác năm khác tháng
        //            || fromdate.getMonth() == 11 && fromdate.getFullYear() + 1 == todate.getFullYear() && difference_In_Days < 30) {

        //            let grid = $("#gridReportDetailtByOneMarket").data("kendoGrid");
        //            grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
        //            grid.dataSource.read();

        //        } else {
        //            $("<div></div>").kendoAlert({
        //                title: "Cảnh báo!",
        //                content: "Bạn chỉ được phép tìm kiếm trong 30 ngày trở lại"
        //            }).data("kendoAlert").open();
        //        }
        //    }
        //});

        // PrintExcel for report day
        $('ul.search-item .print-excel-MarketOne').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

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

                // TypeID = 2 để tạo cột dữ liệu cho thị trường
                window.location = "/ReportDetailtExcelForMarket/CreateExcelMarketForOne/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=2" + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.MarketForOneForMonth = function () {
        
        // PrintExcel for report day
        $('ul.search-item .print-excel-MarketOne-ForMonth').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

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

                // TypeID = 2 để tạo cột dữ liệu cho thị trường
                window.location = "/ReportDetailtExcelForMarket/CreateExcelMarketForOne/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=2" + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    this.GradationCompartion = function () {

        $('ul.search-item .search-grid-gradation-forAll').click(function () {

            //// For month report
            //let urlChart = "/Admin/Report/SearchChartGradationCompare?Gradation=";

            let gradationDicID = $("#gradation").data("kendoComboBox").value();

            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            // Grid hiển thị
            let grid = $("#gridGradationCompare").data("kendoGrid");
            let urlGrid = "/Admin/ReportDetailtByMarket/SearchGridReportForGradation?Gradation=";
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            // Thay đổi title header của grid
            let gradationYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            let gradationLastYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear - 1);
            $($("#gridGradationCompare thead").find("tr:eq(0) th[colspan='4'] .k-link")[0]).html(gradationLastYearText);
            $($("#gridGradationCompare thead").find("tr:eq(0) th[colspan='4'] .k-link")[1]).html(gradationYearText);
            grid.dataSource.read();

            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường
            let chartGradation = $("#chartGradationCompare").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnChartReportForGradation?Gradation=";
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            chartGradation.options.categoryAxis[1].categories = [gradationLastYearText, gradationYearText];
            chartGradation.dataSource.read();

            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPieProportion = $("#ColumnchartForYearPercent").data("kendoChart");
            let urlChartPieProportion = "/Admin/ReportDetailtByMarket/SearchColumnChartReportForGradationPercent?Gradation=";
            chartPieProportion.dataSource.transport.options.read.url = urlChartPieProportion + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartPieProportion.options.title.text = kendo.format("Tỉ trọng từng dịch vụ từng thị trường \n Giai đoạn: {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            chartPieProportion.options.categoryAxis[1].categories = [gradationLastYearText, gradationYearText];
            chartPieProportion.dataSource.read();
        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation-forAll').click(function () {

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportDetailtExcelForMarket/CreateExcelForGradationCompare/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType;
        });
    }

    this.GradationCompartionForOne = function () {

        $('ul.search-item .search-grid-gradation-ForOne').click(function () {
            
            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

            let gradationDicID = $("#gradation").data("kendoComboBox").value();
            let gradationDicText = $("#gradation").data("kendoComboBox").text();

            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            // Grid hiển thị
            let grid = $("#gridGradationCompareForOne").data("kendoGrid");
            let urlGrid = "/Admin/ReportDetailtByMarket/SearchGridReportForGradationForOne?Gradation=";
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            // Thay đổi title header của grid
            let gradationYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            let gradationLastYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear - 1);
            $($("#gridGradationCompareForOne thead").find("tr:eq(0) th[colspan='4'] .k-link")[0]).html(gradationLastYearText);
            $($("#gridGradationCompareForOne thead").find("tr:eq(0) th[colspan='4'] .k-link")[1]).html(gradationYearText);
            grid.dataSource.read();

            // Grid hiển thị
            let gridCompare = $("#gridGradationCompareForOneCompare").data("kendoGrid");
            let urlGridCompare = "/Admin/ReportDetailtByMarket/SearchGridReportForGradationForOneCompare?Gradation=";
            gridCompare.dataSource.transport.options.read.url = urlGridCompare + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            // title cho bảng tăng giảm so với cùng kì năm ngoái
            gradationLastYearText = kendo.format("Tăng giảm so với cùng kì {0}", toYear - 1);
            $($("#gridGradationCompareForOneCompare thead").find("tr:eq(0) th[colspan='4'] .k-link")[0]).html(gradationLastYearText);
            gridCompare.dataSource.read();

            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường
            let chartGradation = $("#chartGradationCompareForOne").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartGradationCompareForOne?Gradation=";
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", $('#categoriesDetaitMarket').data('kendoDropDownList').text(), $("#gradation").data("kendoComboBox").text(), toYear);
            chartGradation.dataSource.read();

            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPieProportion = $("#ColumnchartForYearPercentForOne").data("kendoChart");
            let urlChartPieProportion = "/Admin/ReportDetailtByMarket/SearchColumnChartGradationCompareStackForOne?Gradation=";
            chartPieProportion.dataSource.transport.options.read.url = urlChartPieProportion + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chartPieProportion.options.title.text = kendo.format("Tỉ trọng từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", $('#categoriesDetaitMarket').data('kendoDropDownList').text(), $("#gradation").data("kendoComboBox").text(), toYear);

            chartPieProportion.dataSource.read();

            // Hiển thị theo phần trăm
            // Grid hiển thị
            let gridPercent = $("#gridGradationComparePercent").data("kendoGrid");
            let urlGridPercent = "/Admin/ReportDetailtByMarket/SearchDataDetailtGridGradationComparePercent?Gradation=";
            gridPercent.dataSource.transport.options.read.url = urlGridPercent + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            // Thay đổi giá trị của header
            let titleName = "Lũy kế " + gradationDicText + " " + (toYear - 1).toString();
            $("#gridGradationCompareConvert thead").find("[data-field='LK2'] .k-link").html(titleName);

            titleName = "Lũy kế " + gradationDicText + " " + (toYear).toString();
            $("#gridGradationCompareConvert thead").find("[data-field='LK1'] .k-link").html(titleName);
            gridPercent.dataSource.read();

            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPie = $("#chartGradationPercentYear").data("kendoChart");
            let urlChartPie = "/Admin/ReportDetailtByMarket/SearchDataGradationComparePieYear?Gradation=";
            chartPie.dataSource.transport.options.read.url = urlChartPie + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chartPie.options.title.text = kendo.format("Tỉ trọng từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", $('#categoriesDetaitMarket').data('kendoDropDownList').text(), $("#gradation").data("kendoComboBox").text(), toYear);

            chartPie.dataSource.read();

            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPieLastYear = $("#chartGradationPercentLastYear").data("kendoChart");
            let urlChartPieLastYear = "/Admin/ReportDetailtByMarket/SearchDataGradationComparePieLastYear?Gradation=";
            chartPieLastYear.dataSource.transport.options.read.url = urlChartPieLastYear + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chartPieLastYear.options.title.text = kendo.format("Tỉ trọng từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", $('#categoriesDetaitMarket').data('kendoDropDownList').text(), $("#gradation").data("kendoComboBox").text(), toYear - 1);

            chartPieLastYear.dataSource.read();
        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation-ForOne').click(function () {

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportDetailtExcelForMarket/CreateExcelGradationCompareForOne/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
        });
    }

    this.CompareMonthForAll = function () {

        $('ul.search-item .search-grid-comparemonth-forAll').click(function () {
            
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            let monthYear = kendo.format("Tháng {0}/{1}", month, year);
            let lastYear = kendo.format("Tháng {0}/{1}", month, year - 1);
            let lastMonth = kendo.format("Tháng {0}/{1}", month - 1, year);


            // grid doanh số
            let urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForAll?month=";
            let grid = $("#gridCompareMonthForAll").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;

            $($("#gridCompareMonthForAll thead tr:eq(0) th")[1]).find('.k-link').html(monthYear);
            $($("#gridCompareMonthForAll thead tr:eq(0) th")[2]).find('.k-link').html(lastMonth);
            $($("#gridCompareMonthForAll thead tr:eq(0) th")[3]).find('.k-link').html(lastYear);
            
            grid.dataSource.read();

            // grid doanh số
            urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForAllCompare?month=";
            grid = $("#gridCompareMonthForAllCompare").data("kendoGrid");
            
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            grid.dataSource.read();

            // Biểu đồ cột doanh số theo phần trăm
            let chart = $("#ColumnChartStackCompareMonthForAllPercent").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnChartStackCompareMonthForAll?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chart.options.title.text = kendo.format("Tỉ trọng từng thị trường từng loại dịch vụ \n {0}", monthYear);
            chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số
            chart = $("#chartGradationCompare").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartCompareMonthForAll?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType;
            chart.options.title.text = kendo.format("Doanh số từng thị trường từng loại dịch vụ \n {0}", monthYear);
            chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();
            
        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-grid-comparemonth-forAll').click(function () {
            
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;
            
            window.location = "/ReportDetailtExcelForMarket/CreateExcelCompareForAll/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
        });
    }

    this.CompareMonthForOne = function () {

        $('ul.search-item .search-grid-comparemonth-forOne').click(function () {

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();
            let marketName = $('#categoriesDetaitMarket').data('kendoDropDownList').text();

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            let monthYear = kendo.format("Tháng {0}/{1}", month, year);
            let lastYear = kendo.format("Tháng {0}/{1}", month, year - 1);
            let lastMonth = kendo.format("Tháng {0}/{1}", month - 1, year);


            // grid doanh số
            let urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForOne?month=";
            let grid = $("#gridCompareMonthForOne").data("kendoGrid");
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;

            $($("#gridCompareMonthForOne thead tr:eq(0) th")[1]).find('.k-link').html(monthYear);
            $($("#gridCompareMonthForOne thead tr:eq(0) th")[2]).find('.k-link').html(lastMonth);
            $($("#gridCompareMonthForOne thead tr:eq(0) th")[3]).find('.k-link').html(lastYear);

            grid.dataSource.read();

            // grid doanh số
            urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForOneCompare?month=";
            grid = $("#gridCompareMonthForOneCompare").data("kendoGrid");

            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            grid.dataSource.read();

            // Biểu đồ cột doanh số theo phần trăm
            let chart = $("#chartColumnChartCompareMonthStackForOne").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnChartCompareMonthStackForOne?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", marketName);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số
            chart = $("#chartColumnsChartCompareMonthForOne").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartCompareMonthForOne?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", marketName);
            chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ tròn tỉ trọng của tháng hiện tại
            chart = $("#chartCompareMonthPercentYear").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieYear?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", marketName, monthYear);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ tròn tỉ trọng của tháng trước
            chart = $("#chartCompareMonthPercentLastMonth").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieLastMonth?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", marketName, lastMonth);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ tròn tỉ trọng của cùng kì năm trước
            chart = $("#chartCompareMonthPercentLastYear").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieLastYear?month=";
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", marketName, lastYear);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // grid tỉ trọng với tháng hiện tại so với tháng trước và cùng kì năm trước
            urlGridDS = "/Admin/ReportDetailtByMarket/SearchDataDetailtGridCompareMonthPercent?month=";
            grid = $("#gridCompareMonthPercent").data("kendoGrid");
            $($("#gridCompareMonthPercent thead tr:eq(0) th")[2]).find('.k-link').html(monthYear);
            $($("#gridCompareMonthPercent thead tr:eq(0) th")[3]).find('.k-link').html(lastMonth);
            $($("#gridCompareMonthPercent thead tr:eq(0) th")[4]).find('.k-link').html(lastYear);
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
            grid.dataSource.read();


        });

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-grid-comparemonth-forOne').click(function () {

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            window.location = "/ReportDetailtExcelForMarket/CreateExcelCompareMonthForOne/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
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