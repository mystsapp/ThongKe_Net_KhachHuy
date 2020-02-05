//returns a Date() object in dd/MM/yyyy
$.formattedDate = function (dateToFormat) {
    var dateObject = new Date(dateToFormat);
    var day = dateObject.getDate();
    var month = dateObject.getMonth() + 1;
    var year = dateObject.getFullYear();
    day = day < 10 ? "0" + day : day;
    month = month < 10 ? "0" + month : month;
    var formattedDate = day + "/" + month + "/" + year;
    return formattedDate;
};
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

var khachHuyController = {
    init: function () {
        // khachHuyController.LoadData();
        //var cn = '@Request.RequestContext.HttpContext.Session["chinhanh"]';
        //khachHuyController.loadDdlDaiLyByChiNhanh();
        khachHuyController.loadDdlChiNhanh();
        khachHuyController.registerEvent();
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


        $('#btnExport').off('click').on('click', function () {
            $('#frmSearch').submit();
        });

        $('#btnReset').off('click').on('click', function () {
            khachHuyController.resetForm();
        });

        $('#btnSearch').off('click').on('click', function () {
            if ($('#frmSearch').valid()) {
                khachHuyController.LoadData();
            }
        });

        //$('.btnExportDetail').off('click').on('click', function () {

        //    var tungay = $('#txtTuNgay').val();
        //    var denngay = $('#txtDenNgay').val();
        //    var nhanvien = $(this).data('nhanvien');
        //    var cn = $(this).data('chinhanh');
        //    var khoi = '';
        //    //if (cn == "") {
        //    //    cn = $('#ddlChiNhanh').val();
        //    //    var khoi = $('#ddlKhoi').val();
        //    //} else {
        //    //    var khoi = $('#hidKhoi').data('khoi');
        //    //}
        //    if ($('#hidNhom').val() !== "Users") {
        //        khoi = $('#ddlKhoi').val();
        //    } else {
        //        khoi = $('#hidKhoi').data('khoi');
        //    }

        //    $('#hidTuNgay').val(tungay);
        //    $('#hidDenNgay').val(denngay);
        //    $('#hidNhanVien').val(nhanvien);

        //    $('#hidKhoi').val(khoi);
        //    $('#hidChiNhanh').val(cn);
        //    $('#frmDetail').submit();
        //    //alert(daily);
        //    //quayTheoNgayDiController.ExportDetail();
        //});


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
    //loadDdlDaiLyByChiNhanh: function () {
    //    //var cn = $('#hidCn').data('cn');
    //    var nhom = $('#hidNhom').data('nhom');
    //    $('#ddlDaiLy').html('');
    //    var option = '';
    //    option = '<option value=" ">' + "Tất cả" + '</option>';
    //    $.ajax({
    //        url: '/account/GetDmdailyByNhomChiNhanh',
    //        type: 'GET',
    //        data: {
    //            nhom: nhom
    //        },
    //        dataType: 'json',
    //        success: function (response) {

    //            var data = JSON.parse(response.data);
    //            //console.log(data);

    //            $.each(data, function (i, item) {
    //                option = option + '<option value="' + item.Daily + '">' + item.Daily + '</option>'; //chinhanh1

    //            });
    //            $('#ddlDaiLy').html(option);

    //            //$('#ddlDaiLy option').each(function () {
    //            //    if ($(this).val() == null) {
    //            //        $(this).remove();
    //            //    }
    //            //});
    //            if (nhom != 'Admins')
    //                $("#ddlDaiLy option[value=' ']").remove();

    //        }
    //    });
    //    //homeController.resetForm();
    //    //var id = $('.btn-edit').data('id');


    //},

    LoadData: function (changePageSize) {
        //var name = $('#txtNameS').val();
        //var status = $('#ddlStatusS').val();
        var tungay = $('#txtTuNgay').val();
        var denngay = $('#txtDenNgay').val();
        //var daily = $('#ddlDaiLy').val();
        var hidCn = $('#hidCn').data('cn');
        var chinhanh = $('#ddlChiNhanh').val()
        var khoi = '';

        //var hidCn = '' ? khoi = $('#ddlKhoi').val() : khoi = $('#hidKhoi').data('khoi');
        if (hidCn === "")
            khoi = $('#ddlKhoi').val();
        else {
            khoi = $('#hidKhoi').data('khoi');
        }

        $.ajax({
            url: '/BaoCao/LoadDataKhachHuy',
            type: 'GET',
            data: {
                tungay: tungay,
                denngay: denngay,
                cn: chinhanh,
                khoi: khoi,
                page: homeconfig.pageIndex,
                pageSize: homeconfig.pageSize
            },
            dataType: 'json',
            success: function (response) {
                console.log(response.status);
                if (response.status) {
                    
                    var data = response.data;
                    //var data = JSON.parse(response.data);

                    //alert(data);
                    var html = '';
                    var template = $('#data-template').html();


                    $.each(data, function (i, item) {
                        //usage:
                        //var formattedDate = $.formattedDate(new Date(parseInt(item.ngaysinh.substr(6))));
                        //alert(formattedDate)

                        //var ns = "";
                        //if (item.ngaysinh === null)
                        //    ns = "";
                        //else
                        //    ns = $.formattedDate(new Date(parseInt(item.ngaysinh.substr(6))));

                        var batdau = $.formattedDate(new Date(parseInt(item.batdau.substr(6))));
                        var ketthuc = $.formattedDate(new Date(parseInt(item.ketthuc.substr(6))));
                        var ngayhuyve = $.formattedDate(new Date(parseInt(item.ngayhuyve.substr(6))));

                        html += Mustache.render(template, {
                            stt: item.stt,
                            tenkhach: item.tenkhach,
                            sgtcode: item.sgtcode,
                            vetourid: item.vetourid,
                            tuyentq: item.tuyentq,
                            batdau: batdau,
                            ketthuc: ketthuc,
                            giatour: numeral(item.giatour).format('0,0'),
                            nguoihuyve: item.nguoihuyve,
                            dailyhuyve: item.dailyhuyve,
                            chinhanh: item.chinhanh,
                            ngayhuyve: ngayhuyve

                        });

                    });

                    $('#tblData').html(html);
                    khachHuyController.paging(response.total, function () {
                        khachHuyController.LoadData();
                    }, changePageSize);
                    khachHuyController.registerEvent();
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
khachHuyController.init();