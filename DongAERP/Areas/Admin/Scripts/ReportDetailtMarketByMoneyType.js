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
    this.MarketForOne = function () {

        // For day report
        let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOne?fromDay=";

        $('ul.search-item .search-grid-MarketOne').click(function () {

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

            if (marketID != '') {
                // Chỉ cho phép lấy thời gian trong khoản 30 ngày
                if ((fromdate.getMonth() == todate.getMonth() && fromdate.getFullYear() == todate.getFullYear())
                    // Trường hợp khác tháng cùng năm
                    || (todate.getMonth() == fromdate.getMonth() + 1 && fromdate.getFullYear() == todate.getFullYear() && difference_In_Days < 30)
                    // Trường hợp khác năm khác tháng
                    || fromdate.getMonth() == 11 && fromdate.getFullYear() + 1 == todate.getFullYear() && difference_In_Days < 30) {

                    let grid = $("#gridReportDetailtByOneMarket").data("kendoGrid");
                    grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
                    grid.dataSource.read();

                    //Quy USD
                    let urlGridConvert = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneConvert?fromDay=";
                    let gridConvert = $("#gridReportDetailtByOneMarketConvert").data("kendoGrid");
                    gridConvert.dataSource.transport.options.read.url = urlGridConvert + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + marketID;
                    gridConvert.dataSource.read();

                } else {
                    $("<div></div>").kendoAlert({
                        title: "Cảnh báo!",
                        content: "Bạn chỉ được phép tìm kiếm trong 30 ngày trở lại"
                    }).data("kendoAlert").open();
                }
            }
        });

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

    this.GradationCompartion = function () {

        $('ul.search-item .search-grid-gradation-forAll').click(function () {

            //// For month report
            //let urlChart = "/Admin/Report/SearchChartGradationCompare?Gradation=";

            let gradationDicID = $("#gradation").data("kendoComboBox").value();

            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            // Grid hiển thị
            let grid = $("#gridGradationCompare").data("kendoGrid");
            let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchGridReportForGradation?Gradation=";
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            // Thay đổi title header của grid
            let gradationYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            let gradationLastYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear - 1);
            $($("#gridGradationCompare thead").find("tr:eq(0) th[data-field !='MarketName'] .k-link")[0]).html(gradationLastYearText);
            $($("#gridGradationCompare thead").find("tr:eq(0) th[data-field !='MarketName'] .k-link")[1]).html(gradationYearText);
            grid.dataSource.read();

            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường
            let chartGradation = $("#chartGradationCompare").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnChartReportForGradation?Gradation=";
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            chartGradation.options.categoryAxis[1].categories = [gradationLastYearText, gradationYearText];
            chartGradation.dataSource.read();

            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPieProportion = $("#ColumnchartForYearPercent").data("kendoChart");
            let urlChartPieProportion = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnChartReportForGradationPercent?Gradation=";
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

    /// Call select item khi chọn item từ thị trường con của thị trường Châu Á
    this.onSelectGradationMarketID = function (e) {
        debugger;
        let dataItem = null;
        // Khi chọn dữ liệu từ các thị trường con của Thị trường Châu Á
        if (e.item) {

            dataItem = e.dataItem;
        }

        if (dataItem != null) {

            // Get mã thị trường
            let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();
            let marketName = $('#categoriesDetaitMarket').data('kendoDropDownList').text();

            let gradationDicID = $("#gradation").data("kendoComboBox").value();
            let gradationDicText = $("#gradation").data("kendoComboBox").text();

            let marketAsianDetailtText = $('#dropdownlistMarket').data('kendoDropDownList').text();

            let toYear = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            // Thay đổi title header của grid
            let gradationYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear);
            let gradationLastYearText = kendo.format("Lũy kế {0} {1}", $("#gradation").data("kendoComboBox").text(), toYear - 1);

            // Grid hiển thị
            let grid = $("#gridGradationCompareForOne").data("kendoGrid");
            let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchGridReportForGradationForOne?Gradation=";
            grid.dataSource.transport.options.read.data = '';
            grid.dataSource.transport.options.read.url = urlGrid + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            
            //$($("#gridGradationCompareForOne thead").find("tr:eq(0) th[data-field != 'PartnerName'] .k-link")[0]).html(gradationLastYearText);
            //$($("#gridGradationCompareForOne thead").find("tr:eq(0) th[data-field != 'PartnerName'] .k-link")[1]).html(gradationYearText);
            grid.dataSource.read();

            // Xóa dòng tổng với trường hợp chỉ có 1 thị trường
            if (dataItem.Value != '005') {
                grid.bind("dataBound", function () {
                    $('#gridGradationCompareForOne tr.k-group-footer').css('display', 'none');
                });
            } else {
                grid.bind("dataBound", function () {
                    $('#gridGradationCompareForOne tr.k-group-footer').css('display', '');
                });
            }

            // Grid hiển thị
            let gridCompare = $("#gridGradationCompareForOneCompare").data("kendoGrid");
            let urlGridCompare = "/Admin/ReportDetailtMarketByMoneyType/SearchGridReportForGradationForOneCompare?Gradation=";
            gridCompare.dataSource.transport.options.read.data = '';
            gridCompare.dataSource.transport.options.read.url = urlGridCompare + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            // title cho bảng tăng giảm so với cùng kì năm ngoái
            //gradationLastYearText = kendo.format("Tăng giảm so với cùng kì {0}", toYear - 1);
            //$($("#gridGradationCompareForOneCompare thead").find("tr:eq(0) th[data-field != 'PartnerName'] .k-link")[0]).html(gradationLastYearText);
            gridCompare.dataSource.read();
            
            // Xóa dòng tổng với trường hợp chỉ có 1 thị trường
            if (dataItem.Value != '005') {
                grid.bind("dataBound", function () {
                    $('#gridGradationCompareForOneCompare tr.k-group-footer').css('display', 'none');
                });
            } else {
                grid.bind("dataBound", function () {
                    $('#gridGradationCompareForOneCompare tr.k-group-footer').css('display', '');
                });
            }
            
            // Biểu đồ cột tỉ trọng từng dịch vụ từng thị trường
            let chartPieProportion = $("#ColumnchartForYearPercentForOne").data("kendoChart");
            let urlChartPieProportion = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnChartGradationCompareStackForOne?gradation=";
            chartPieProportion.dataSource.transport.options.read.data = '';
            chartPieProportion.dataSource.transport.options.read.url = urlChartPieProportion + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartPieProportion.options.title.text = kendo.format("Tỉ trọng từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);

            chartPieProportion.options.categoryAxis[1].categories = ["AUD", "CAD", "EUR", "GBP", "USD", "VND"];
            chartPieProportion.dataSource.read();


            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - VND
            let chartGradation = $("#chartGradationCompareForOneVND").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneVND?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);

            //chartGradation.options.categoryAxis[1].categories = [gradationLastYearText, gradationYearText];
            chartGradation.dataSource.read();

            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - USD
            chartGradation = $("#chartGradationCompareForOneUSD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneUSD?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);
            
            chartGradation.dataSource.read();


            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - EUR
            chartGradation = $("#chartGradationCompareForOneEUR").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneEUR?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);
            
            chartGradation.dataSource.read();


            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - CAD
            chartGradation = $("#chartGradationCompareForOneCAD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneCAD?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);
            
            chartGradation.dataSource.read();


            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - AUD
            chartGradation = $("#chartGradationCompareForOneAUD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneAUD?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);
            
            chartGradation.dataSource.read();


            // Biểu đồ cột hiển thị từng dịch vụ từng thị trường - GBP
            chartGradation = $("#chartGradationCompareForOneGBP").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartGradationCompareForOneGBP?Gradation=";
            chartGradation.dataSource.transport.options.read.data = '';
            chartGradation.dataSource.transport.options.read.url = urlChartDS + gradationDicID + "&year=" + toYear + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chartGradation.options.title.text = kendo.format("Doanh số từng dịch vụ từng thị trường \n Thị trường: {0} \n Giai đoạn: {1} {2}", dataItem.Text, $("#gradation").data("kendoComboBox").text(), toYear);
            
            chartGradation.dataSource.read();

        }
    };

    this.CompareMonthForAll = function () {
        
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

    /// Call select item khi chọn item từ thị trường con của thị trường Châu Á
    this.onSelectMarketID = function (e) {
        debugger;
        let dataItem = null;
        // Get mã thị trường
        let marketID = $('#categoriesDetaitMarket').data('kendoDropDownList').value();
        let marketName = $('#categoriesDetaitMarket').data('kendoDropDownList').text();

        let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
        let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

        let monthYear = kendo.format("Tháng {0}/{1}", month, year);
        let lastYear = kendo.format("Tháng {0}/{1}", month, year - 1);
        let lastMonth = kendo.format("Tháng {0}/{1}", month - 1, year);

        // Khi chọn dữ liệu từ các thị trường con của Thị trường Châu Á
        if (e.item) {

            dataItem = e.dataItem;
        }

        if (dataItem != null) {

            // grid bảng số liệu
            let urlGridDS = "/Admin/ReportDetailtMarketByMoneyType/SearchReportDetailtCompareMonthForOne?month=";
            let grid = $("#gridCompareMonthForOne").data("kendoGrid");
            grid.dataSource.transport.options.read.data = '';
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            grid.dataSource.read();

            // Xóa dòng tổng với trường hợp chỉ có 1 thị trường
            if (dataItem.Value != '005') {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthForOne tr.k-group-footer').css('display', 'none');
                });
            } else {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthForOne tr.k-group-footer').css('display', '');
                });
            }

            // grid cho bảng số liệu tăng giảm
            urlGridDS = "/Admin/ReportDetailtMarketByMoneyType/SearchReportDetailtCompareMonthForOneCompare?month=";
            grid = $("#gridCompareMonthForOneCompare").data("kendoGrid");
            grid.dataSource.transport.options.read.data = '';
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            grid.dataSource.read();

            // Xóa dòng tổng với trường hợp chỉ có 1 thị trường
            if (dataItem.Value != '005') {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthForOneCompare tr.k-group-footer').css('display', 'none');
                });
            } else {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthForOneCompare tr.k-group-footer').css('display', '');
                });
            }
            
            // Biểu đồ cột doanh số - VND
            chart = $("#chartColumnsChartCompareMonthForOneVND").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneVND?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số - USD
            chart = $("#chartColumnsChartCompareMonthForOneUSD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneUSD?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số - EUR
            chart = $("#chartColumnsChartCompareMonthForOneEUR").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneEUR?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số - CAD
            chart = $("#chartColumnsChartCompareMonthForOneCAD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneCAD?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số - AUD
            chart = $("#chartColumnsChartCompareMonthForOneAUD").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneAUD?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số - GBP
            chart = $("#chartColumnsChartCompareMonthForOneGBP").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnsChartCompareMonthForOneGBP?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // Biểu đồ cột doanh số theo phần trăm
            chart = $("#chartColumnChartCompareMonthStackForOne").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtMarketByMoneyType/SearchColumnChartCompareMonthStackForOne?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", + dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Xóa dòng tổng với trường hợp chỉ có 1 thị trường
            if (dataItem.Value != '005') {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthPercent tr.k-group-footer').css('display', 'none');
                });
            } else {
                grid.bind("dataBound", function () {
                    $('#gridCompareMonthPercent tr.k-group-footer').css('display', '');
                });
            }

        }
    };

    /// Call select item khi chọn item từ thị trường con của thị trường Châu Á
    this.onSelectMarketForOne = function (e) {
        debugger;
        let dataItem = null;
        // Khi chọn dữ liệu từ các thị trường con của Thị trường Châu Á
        if (e.item) {

            dataItem = e.dataItem;
        }

        if (dataItem != null) {

            let fromDate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số ngày chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);
            // Grid hiển thị

            // Chỉ cho phép lấy thời gian trong khoản 30 ngày
            if ((fromDate.getMonth() == toDate.getMonth() && fromDate.getFullYear() == toDate.getFullYear())
                // Trường hợp khác tháng cùng năm
                || (toDate.getMonth() == fromDate.getMonth() + 1 && fromDate.getFullYear() == toDate.getFullYear() && difference_In_Days < 30)
                // Trường hợp khác năm khác tháng
                || fromDate.getMonth() == 11 && fromDate.getFullYear() + 1 == toDate.getFullYear() && difference_In_Days < 30) {

                let grid = $("#gridReportDetailtByOneMarket").data("kendoGrid");
                let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOne?fromDay=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

                grid = $("#gridReportDetailtByOneMarketConvert").data("kendoGrid");
                urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneConvert?fromDay=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromDateConvert + "&toDay=" + toDateConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 30 ngày trở lại"
                }).data("kendoAlert").open();
            }
        }
    };

    /// Call select item khi chọn item từ thị trường con của thị trường Châu Á
    this.onSelectMarketForOneForMonth = function (e) {
        debugger;
        let dataItem = null;
        // Khi chọn dữ liệu từ các thị trường con của Thị trường Châu Á
        if (e.item) {

            dataItem = e.dataItem;
        }

        if (dataItem != null) {

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

                let grid = $("#gridReportDetailtByOneMarketForMonth").data("kendoGrid");
                let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneForMonth?fromDate=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

                grid = $("#gridReportDetailtByOneMarketForMonthConvert").data("kendoGrid");
                urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneForMonthConvert?fromDate=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromMonthConvert + "&toDate=" + toMonthConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        }
    };

    /// Call select item khi chọn item từ thị trường con của thị trường Châu Á
    this.onSelectMarketForOneForYear = function (e) {
        debugger;
        let dataItem = null;
        // Khi chọn dữ liệu từ các thị trường con của Thị trường Châu Á
        if (e.item) {

            dataItem = e.dataItem;
        }

        if (dataItem != null) {

            let fromYear = $('#FromYear').data('kendoDatePicker').value();
            let fromYearConvert = kendo.toString(fromYear, "MM/dd/yyyy");
            let toYear = $('#ToYear').data('kendoDatePicker').value();
            let toYearConvert = kendo.toString(toYear, "MM/dd/yyyy");

            // Chỉ show dc 5 năm
            if (toYear.getFullYear() - fromYear.getFullYear() <= 4) {

                let grid = $("#gridReportDetailtByOneMarketForYear").data("kendoGrid");
                let urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneForYear?fromDate=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

                grid = $("#gridReportDetailtByOneMarketForYearConvert").data("kendoGrid");
                urlGrid = "/Admin/ReportDetailtMarketByMoneyType/SearchMarketForOneForYearConvert?fromDate=";
                grid.dataSource.transport.options.read.data = '';
                grid.dataSource.transport.options.read.url = urlGrid + fromYearConvert + "&toDate=" + toYearConvert + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;

                grid.dataSource.read();

            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 5 năm trở lại"
                }).data("kendoAlert").open();
            }
        }
    };
}

// Loading
function displayLoading(target) {

    var element = $(target);
    kendo.ui.progress(element, true);
    setTimeout(function () {
        kendo.ui.progress(element, false);
    }, 2000);
}