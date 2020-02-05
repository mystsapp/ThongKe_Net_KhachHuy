
$.validator.addMethod("dateFormat",
    function (value, element) {
        var check = false;
        var re = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
        if (re.test(value)) {
            var adata = value.split('/');
            var dd = parseInt(adata[0], 10);
            var mm = parseInt(adata[1], 10);
            var yyyy = parseInt(adata[2], 10);
            var xdata = new Date(yyyy, mm - 1, dd);
            if ((xdata.getFullYear() === yyyy) && (xdata.getMonth() === mm - 1) && (xdata.getDate() === dd)) {
                check = true;
            }
            else {
                check = false;
            }
        } else {
            check = false;
        }
        return this.optional(element) || check;
    },
    "Chưa đúng định dạng dd/mm/yyyy.");

var homeconfig = {
    pageSize: 15,
    pageIndex: 1
};

var kinhDoanhOnlineController = {
    init: function () {
        // kinhDoanhOnlineController.LoadData();
        //var cn = '@Request.RequestContext.HttpContext.Session["chinhanh"]';
        //kinhDoanhOnlineController.loadDdlDaiLyByChiNhanh();
        kinhDoanhOnlineController.registerEvent();
    },

    registerEvent: function () {

        //$('.modal-dialog').draggable();

        $('#frmSearch').validate({
            rules: {
                tungay: {
                    required: true,
                    dateFormat: true
                    //date: true
                },
                denngay: {
                    required: true,
                    dateFormat: true
                    //date: true
                }
            },
            messages: {
                tungay: {
                    required: "Trường này không được để trống."
                    //date: "Chưa đúng định dạng."
                },
                denngay: {
                    required: "Trường này không được để trống."
                    //number: "password phải là số",
                    //date: "Chưa đúng định dạng."
                }
            }
        });


        //$('#btnExport').off('click').on('click', function () {
        //    $('#frmSearch').submit();
        //});

        $('#btnReset').off('click').on('click', function () {
            kinhDoanhOnlineController.resetForm();
        });

        $('#btnSearch').off('click').on('click', function () {
            if ($('#frmSearch').valid()) {
                kinhDoanhOnlineController.LoadData();
            }
        });

        $('.btnExportDetail').off('click').on('click', function () {

            var tungay = $('#txtTuNgay').val();
            var denngay = $('#txtDenNgay').val();
            var cn = $(this).data('chinhanh');
            var khoi = '';

            if ($('#hidNhom').val() !== "Users") {
                khoi = $('#ddlKhoi').val();
            } else {
                khoi = $('#hidKhoi').data('khoi');
            }

            $('#hidTuNgay').val(tungay);
            $('#hidDenNgay').val(denngay);
            $('#hidChiNhanh').val(cn);
            $('#hidKhoi').val(khoi);

            $('#frmDetail').submit();

        });


        $("#txtTuNgay, #txtDenNgay").datepicker({
            changeMonth: true,
            changeYear: true,
            dateFormat: "dd/mm/yy"

        });
    },
    resetForm: function () {
        $('#txtTuNgay').val('');
        $('#txtDenNgay').val('');
    },

    LoadData: function (changePageSize) {
        var tungay = $('#txtTuNgay').val();
        var denngay = $('#txtDenNgay').val();
        var hidCn = $('#hidCn').data('cn');
        var khoi = '';

        if (hidCn === "")
            khoi = $('#ddlKhoi').val();
        else {
            khoi = $('#hidKhoi').data('khoi');
        }

        $.ajax({
            url: '/BaoCao/LoadDataKinhDoanhOnline',
            type: 'GET',
            data: {
                tungay: tungay,
                denngay: denngay,
                khoi: khoi,
                page: homeconfig.pageIndex,
                pageSize: homeconfig.pageSize
            },
            dataType: 'json',
            success: function (response) {
                console.log(response.status);
                if (response.status) {
                    console.log(response.data);
                    var data = response.data;
                    //var data = JSON.parse(response.data);

                    //alert(data);
                    var html = '';
                    var template = $('#data-template').html();


                    $.each(data, function (i, item) {

                        html += Mustache.render(template, {
                            chinhanh: item.chinhanh,
                            taove: item.taove,
                            tuhuy: item.tuhuy,
                            chuaxuatve: item.chuaxuatve,
                            thanhcong: item.thanhcong,
                            huy: item.huy
                        });

                    });

                    $('#tblData').html(html);
                    kinhDoanhOnlineController.paging(response.total, function () {
                        kinhDoanhOnlineController.LoadData();
                    }, changePageSize);
                    kinhDoanhOnlineController.registerEvent();
                }
            }
        })
    },

    paging: function (totalRow, callback, changePageSize) {
        var totalPage = Math.ceil(totalRow / homeconfig.pageSize);//lam tron len

        //unbind pagination if it existed or click change size ==> reset
        if (('#pagination a').length === 0 || changePageSize === true) {
            $('#pagination').empty();
            $('#pagination').removeData('twbsPagination');
            $('#pagination').unbind("page");
        }

        $('#pagination').twbsPagination({
            totalPages: totalPage,
            first: "Đầu",
            next: "Tiếp",
            last: "Cuối",
            prev: "trước",
            visiblePages: 10, // tong so trang hien thi , ...12345678910...
            onPageClick: function (event, page) {
                homeconfig.pageIndex = page;//khi chuyen trang thi set lai page index
                setTimeout(callback, 200);
            }
        });
    }

}
kinhDoanhOnlineController.init();