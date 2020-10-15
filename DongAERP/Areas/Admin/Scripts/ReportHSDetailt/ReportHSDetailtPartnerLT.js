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

    // Get value từ localStogra
    valueReportType = localStorage.reportTypeLS;

    //DayReport.GetAdditionalData();
    Report.PartnerForTotal();
    Report.PartnerForTotalForMonth();
    Report.PartnerForTotalForYear();
    
    Report.PartnerForOne();
    Report.PartnerForOneForMonth();
    Report.PartnerForOneForYear();
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
    this.PartnerForTotal = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll').click(function () {

            // loadding
            displayLoading(document.body);

            let fromDate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số ngày chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

            // Chỉ cho phép lấy thời gian trong khoản 30 ngày
            if ((fromDate.getMonth() == toDate.getMonth() && fromDate.getFullYear() == toDate.getFullYear())
                // Trường hợp khác tháng cùng năm
                || (toDate.getMonth() == fromDate.getMonth() + 1 && fromDate.getFullYear() == toDate.getFullYear() && difference_In_Days < 30)
                // Trường hợp khác năm khác tháng
                || fromDate.getMonth() == 11 && fromDate.getFullYear() + 1 == toDate.getFullYear() && difference_In_Days < 30) {

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYear/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType;
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
    this.PartnerForTotalForMonth = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll-ofMonth').click(function () {

            // loadding
            displayLoading(document.body);

            let fromDate = $('#FromMonth').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToMonth').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Chỉ cho phép chọn trong 12 tháng
            if ((fromDate.getFullYear() == toDate.getFullYear() && fromDate.getFullYear() == toDate.getFullYear())
                || (toDate.getFullYear() == fromDate.getFullYear() + 1 && difference_In_Days < 12)) {

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYear/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=2" + "&reportTypeID=" + valueReportType;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.PartnerForTotalForYear = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-forAll-ofYear').click(function () {

            // loadding
            displayLoading(document.body);

            let fromDate = $('#FromYear').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToYear').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Chỉ show dc 5 năm
            if (toDate.getFullYear() - fromDate.getFullYear() > 4) {

                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 5 năm trở lại"
                }).data("kendoAlert").open();
            } else {

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYear/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=3" + "&reportTypeID=" + valueReportType;
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.PartnerForOne = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-PartnerOne').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let partnerID = $('#categoriesDetaitPartner').data('kendoDropDownList').value();

            let fromDate = $('#FromDay').data('kendoDatePicker').value();
            let fromDateConvert = kendo.toString(fromDate, "MM/dd/yyyy");
            let toDate = $('#ToDay').data('kendoDatePicker').value();
            let toDateConvert = kendo.toString(toDate, "MM/dd/yyyy");

            // Tính số ngày chênh lệch
            let difference_In_Time = toDate.getTime() - fromDate.getTime();
            let difference_In_Days = difference_In_Time / (1000 * 3600 * 24);

            // Chỉ cho phép lấy thời gian trong khoản 30 ngày
            if ((fromDate.getMonth() == toDate.getMonth() && fromDate.getFullYear() == toDate.getFullYear())
                // Trường hợp khác tháng cùng năm
                || (toDate.getMonth() == fromDate.getMonth() + 1 && fromDate.getFullYear() == toDate.getFullYear() && difference_In_Days < 30)
                // Trường hợp khác năm khác tháng
                || fromDate.getMonth() == 11 && fromDate.getFullYear() + 1 == toDate.getFullYear() && difference_In_Days < 30) {

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYearForOne/?fromDate=" + fromDateConvert + "&toDate=" + toDateConvert + "&typeID=1" + "&reportTypeID=" + valueReportType + "&partnerID=" + partnerID;
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
    this.PartnerForOneForMonth = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-PartnerOne-ForMonth').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let partnerID = $('#categoriesDetaitPartner').data('kendoDropDownList').value();

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

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYearForOne/?fromDate=" + fromMonthConvert + "&toDate=" + toMonthConvert + "&typeID=2" + "&reportTypeID=" + valueReportType + "&partnerID=" + partnerID;
            } else {
                $("<div></div>").kendoAlert({
                    title: "Cảnh báo!",
                    content: "Bạn chỉ được phép in báo cáo trong 12 tháng trở lại"
                }).data("kendoAlert").open();
            }
        });
    }

    /**
     * Xử lý sự kiện search cho màn hình chi tiết theo ngày cho tất cả
     * @param {any} e
     * @since[Trường Lãm] created on [15/06/2020]
     */
    this.PartnerForOneForYear = function () {

        // PrintExcel for report day
        $('ul.search-item .print-excel-PartnerOne-ForYear').click(function () {

            // loadding
            displayLoading(document.body);

            // Get mã thị trường
            let partnerID = $('#categoriesDetaitPartner').data('kendoDropDownList').value();

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

                window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForDayMonthYearForOne/?fromDate=" + fromYearConvert + "&toDate=" + toYearConvert + "&typeID=3" + "&reportTypeID=" + valueReportType + "&partnerID=" + partnerID;
            }
        });
    }

    this.GradationCompartion = function () {

        $('ul.search-item .print-excel-gradation-forAll').click(function () {

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForGradationCompare/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType;
        });
    }

    this.GradationCompartionForOne = function () {

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-excel-gradation-ForOne').click(function () {

            // Get mã đối tác
            let partnerID = $('#categoriesDetaitPartner').data('kendoDropDownList').value();

            let gradationID = $("#gradation").data("kendoComboBox").value();
            let year = $('#ToYear').data('kendoDatePicker').value().getFullYear();

            window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForGradationCompareForOne/?gradationID=" + gradationID + "&year=" + year + "&reportTypeID=" + valueReportType + "&partnerID=" + partnerID;
        });
    }

    this.CompareMonthForAll = function () {

        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-grid-comparemonth-forAll').click(function () {
            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelForCompareForMonth/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType;
        });
    }

    this.CompareMonthForOne = function (e) {


        // PrintExcel for report month
        // Trường hợp in Excel cho loại hình dịch vụ so sánh theo giao đoạn có type.
        // Trong đó type = 1: Loại hình dịch vụ so sánh theo giai đoạn
        // Trong đó type = 2: Loại hình dịch vụ so sánh theo tháng
        $('ul.search-item .print-grid-comparemonth-forOne').click(function () {

            // Get mã đối tác
            let partnerID = $('#categoriesDetaitPartner').data('kendoDropDownList').value();

            let year = $('#FromMonth').data('kendoDatePicker').value().getFullYear();
            let month = $('#FromMonth').data('kendoDatePicker').value().getMonth() + 1;

            window.location = "/ReportHSDetailtPartnerLTExcel/CreateExcelCompareMonthForOne/?year=" + year + "&month=" + month + "&reportTypeID=" + valueReportType + "&partnerID=" + partnerID;
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
            let urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForOne?month=";
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
            urlGridDS = "/Admin/ReportDetailtByMarket/SearchReportDetailtCompareMonthForOneCompare?month=";
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


            // Biểu đồ cột doanh số theo phần trăm
            let chart = $("#chartColumnChartCompareMonthStackForOne").data("kendoChart");
            let urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnChartCompareMonthStackForOne?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // Biểu đồ cột doanh số cho Doanh số chuyển khoản
            chart = $("#chartColumnsChartCompareMonthForOneForDSCK").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartCompareMonthForOneDSCK?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số cho Doanh số chi nhà
            chart = $("#chartColumnsChartCompareMonthForOneForDSChiNha").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartCompareMonthForOneDSChiNha?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();

            // Biểu đồ cột doanh số cho Doanh số chi Quầy
            chart = $("#chartColumnsChartCompareMonthForOneForDSChiQuay").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchColumnsChartCompareMonthForOneDSChiQuay?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Doanh số các đối tác thị trường {0}", dataItem.Text);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // Biểu đồ tròn tỉ trọng của cùng kì năm trước
            chart = $("#chartCompareMonthPercentLastYear").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieLastYear?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", dataItem.Text, lastYear);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // Biểu đồ tròn tỉ trọng của tháng trước
            chart = $("#chartCompareMonthPercentLastMonth").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieLastMonth?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", dataItem.Text, lastMonth);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // Biểu đồ tròn tỉ trọng của tháng hiện tại
            chart = $("#chartCompareMonthPercentYear").data("kendoChart");
            urlChartDS = "/Admin/ReportDetailtByMarket/SearchDataCompareMonthPieYear?month=";
            chart.dataSource.transport.options.read.data = '';
            chart.dataSource.transport.options.read.url = urlChartDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            chart.options.title.text = kendo.format("Tỉ trong các đối tác thị trường {0} \n {1}", dataItem.Text, monthYear);
            //chart.options.categoryAxis[1].categories = [lastYear, lastMonth, monthYear];
            chart.dataSource.read();


            // grid tỉ trọng với tháng hiện tại so với tháng trước và cùng kì năm trước
            urlGridDS = "/Admin/ReportDetailtByMarket/SearchDataDetailtGridCompareMonthPercent?month=";
            grid = $("#gridCompareMonthPercent").data("kendoGrid");
            grid.dataSource.transport.options.read.data = '';
            //$($("#gridCompareMonthPercent thead tr:eq(0) th")[2]).find('.k-link').html(monthYear);
            //$($("#gridCompareMonthPercent thead tr:eq(0) th")[3]).find('.k-link').html(lastMonth);
            //$($("#gridCompareMonthPercent thead tr:eq(0) th")[4]).find('.k-link').html(lastYear);
            grid.dataSource.transport.options.read.url = urlGridDS + month + "&year=" + year + "&reportTypeID=" + valueReportType + "&marketID=" + dataItem.Value;
            grid.dataSource.read();

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

}

// Loading
function displayLoading(target) {

    var element = $(target);
    kendo.ui.progress(element, true);
    setTimeout(function () {
        kendo.ui.progress(element, false);
    }, 2000);
}