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

var saleTheoNgayDiController = {
    init: function () {
        // saleTheoNgayDiController.LoadData();
        //saleTheoNgayDiController.loadDdlDaiLy();
        saleTheoNgayDiController.loadDdlChiNhanh();
        saleTheoNgayDiController.registerEvent();
    },

    registerEvent: function () {

        //$('.modal-dialog').draggable();

        $('#frmSearch').validate({
            rules: {
                tungay: {
                    required: true,
                    //date: true
                    dateFormat: true
                },
                denngay: {
                    required: true,
                    //date: true
                    dateFormat: true
                }
            },
            messages: {
                tungay: {
                    required: "Trường này không được để trống.",
                    //date: "Chưa đúng định dạng."
                },
                denngay: {
                    required: "Trường này không được để trống.",
                    //number: "password phải là số",
                    //date: "Chưa đúng định dạng."
                }
            }
        });


        $('#btnExport').off('click').on('click', function () {
            $('#frmSearch').submit();
        });

        $('#btnReset').off('click').on('click', function () {
            saleTheoNgayDiController.resetForm();
        });

        $('#btnSearch').off('click').on('click', function () {
            if ($('#frmSearch').valid()) {
                saleTheoNgayDiController.LoadData();
            }
        });

        $('.btnExportDetail').off('click').on('click', function () {

            var tungay = $('#txtTuNgay').val();
            var denngay = $('#txtDenNgay').val();
            var cn = $(this).data('chinhanh');
            var nhanvien = $(this).data('nhanvien');

            //if (cn == "") {
            //    cn = $('#ddlChiNhanh').val();
            //    var khoi = $('#ddlKhoi').val();
            //} else {
            //    var khoi = $('#hidKhoi').data('khoi');
            //}

            if ($('#hidNhom').val() != "Users") {
                var khoi = $('#ddlKhoi').val();
            } else {
                var khoi = $('#hidKhoi').data('khoi');
            }

            $('#hidTuNgay').val(tungay);
            $('#hidDenNgay').val(denngay);
            $('#hidNhanVien').val(nhanvien);
            $('#hidChiNhanh').val(cn);
            $('#hidKhoi').val(khoi);

            $('#frmDetail').submit();
            //alert(daily);
            //quayTheoNgayDiController.ExportDetail();
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

    loadDdlChiNhanh: function () {
        $('#ddlChiNhanh').html('');
        var option = '';
        // option = option + '<option value=select>Select</option>';
        var nhom = $('#hidNhom').data('nhom');
        var chinhanh = '';
        if (nhom === 'Users')
            chinhanh = $('#hidCn').data('cn');

        $.ajax({
            url: '/account/GetAllChiNhanhByNhom',
            type: 'GET',
            data: {
                nhom: nhom,
                chinhanh: chinhanh
            },
            dataType: 'json',
            success: function (response) {

                var data = JSON.parse(response.data);
                $('#ddlChiNhanh').html('');

                //$.each(data, function (i, item) {
                //    option = option + '<option value="' + item.chinhanh1 + '">' + item.chinhanh1 + '</option>'; //chinhanh1

                //});

                for (var i = 0; i < data.length; i++) {
                    // set the key/property (input element) for your object
                    var ele = data[i];
                    //console.log(ele);
                    option = option + '<option value="' + ele + '">' + ele + '</option>'; //chinhanh1
                    // add the property to the object and set the value
                    //params[ele] = $('#' + ele).val();
                }

                $('#ddlChiNhanh').html(option);

            }
        });

    },
    //loadDdlDaiLy: function () {

    //    $('#ddlDaiLy').html('');
    //    //var cn = $('#hidCn').data('cn');
    //    var nhom = $('#hidNhom').data('nhom');
    //    //alert(cn);
    //    var option = '';
    //    option = '<option value=" ">' + "Tất cả" + '</option>';
    //    // option = option + '<option value=select>Select</option>';

    //    $.ajax({
    //        url: '/account/GetDmdailyByNhomChiNhanh',
    //        type: 'GET',
    //        data: {
    //            nhom: nhom
    //        },
    //        dataType: 'json',
    //        success: function (response) {
    //            //if (response.length > 0) {
    //            //var data = JSON.stringify(response.data);
    //            var data = JSON.parse(response.data);
    //            console.log(data);

    //            $.each(data, function (i, item) {
    //                option = option + '<option value="' + item.Daily + '">' + item.Daily + '</option>'; //chinhanh1

    //            });
    //            $('#ddlDaiLy').html(option);

    //            if (nhom != 'Admins')
    //                $("#ddlDaiLy option[value=' ']").remove();

    //        }
    //    });


    //},

    LoadData: function (changePageSize) {
        var tungay = $('#txtTuNgay').val();
        var denngay = $('#txtDenNgay').val();
        //var daily = $('#ddlDaiLy').val();
        var cn = $('#hidCn').data('cn');
        if (cn == "") {
            var khoi = $('#ddlKhoi').val();
            cn = $('#ddlChiNhanh').val();
        } else {
            var khoi = $('#hidKhoi').data('khoi');
        }

        $.ajax({
            url: '/BaoCao/LoadDataSaleTheoNgayDi',
            type: 'GET',
            data: {
                tungay: tungay,
                denngay: denngay,
                //daily: daily,
                chinhanh: cn,
                khoi: khoi,
                page: homeconfig.pageIndex,
                pageSize: homeconfig.pageSize
            },
            dataType: 'json',
            success: function (response) {
                //console.log(response.data);
                if (response.status) {
                    //console.log(response.data);
                    var data = response.data;
                    //var data = JSON.parse(response.data);

                    //alert(data);
                    var html = '';
                    var template = $('#data-template').html();

                    $.each(data, function (i, item) {

                        html += Mustache.render(template, {
                            stt: item.stt,
                            nguoixuatve: item.nguoixuatve,
                            chinhanh: item.chinhanh,
                            doanhso: numeral(item.doanhso).format('0,0'),
                            thucthu: numeral(item.thucthu).format('0,0')
                        });

                    })

                    $('#tblData').html(html);
                    saleTheoNgayDiController.paging(response.total, function () {
                        saleTheoNgayDiController.LoadData();
                    }, changePageSize);
                    saleTheoNgayDiController.registerEvent();
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
saleTheoNgayDiController.init();