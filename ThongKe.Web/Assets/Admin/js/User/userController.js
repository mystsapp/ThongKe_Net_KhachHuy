var homeconfig = {
    pageSize: 10,
    pageIndex: 1
}
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


$.stringToDate = function (_date, _format, _delimiter) {
    var formatLowerCase = _format.toLowerCase();
    var formatItems = formatLowerCase.split(_delimiter);
    var dateItems = _date.split(_delimiter);
    var monthIndex = formatItems.indexOf("mm");
    var dayIndex = formatItems.indexOf("dd");
    var yearIndex = formatItems.indexOf("yyyy");
    var month = parseInt(dateItems[monthIndex]);
    month -= 1;
    var formatedDate = new Date(dateItems[yearIndex], month, dateItems[dayIndex]);
    return formatedDate;
}

var userController = {
    init: function () {
        userController.LoadData();

    },

    registerEvent: function () {

        //$('.modal-dialog').draggable();

        $('#frmSaveData').validate({
            rules: {
                username: {
                    required: true,
                    minlength: 3
                },
                password: {
                    //required: true,
                    minlength: 3
                },
                hoten: {
                    required: true
                },
                role: {
                    required: true
                },
                khoi: {
                    required: true
                }
            },
            messages: {
                username: {
                    required: "Trường này không được để trống.",
                    minlength: "Username phải có ít nhất 3 ký tự"
                },
                password: {
                    required: "Trường này không được để trống.",
                    //number: "password phải là số",
                    minlength: "Password phải có ít nhất 3 ký tự"
                },
                hoten: {
                    required: "Trường này không được để trống."
                },
                role: {
                    required: "Trường này không được để trống."
                },
                khoi: {
                    required: "Trường này không được để trống."
                }
            }
        });


        $('#btnAddNew').off('click').on('click', function () {
            $('#modalAddUpdate').modal('show');
            $('#txtUsername').prop("disabled", false);

            userController.resetForm();
            userController.loadDdlChiNhanh();
            userController.loadDdlDaiLy();
        });

        $('#ddlNhom').off('change').on('change', function () {
            //var optionValue = $(this).val();
            //userController.loadDdlDaiLyByChiNhanh(optionValue);
            var optionValue = $(this).val();
            if (optionValue === 'Users') {
                $('#ddlChiNhanh').prop("disabled", false);
                $('#ddlDaiLy').prop("disabled", false);
                //console.log('user.');
            }
            else {
                $('#ddlChiNhanh').prop("disabled", true);
                $('#ddlDaiLy').prop("disabled", true);
                //console.log('admin.');
            }

        });

        $('#ddlChiNhanh').off('change').on('change', function () {
            var optionValue = $(this).val();
            userController.loadDdlDaiLyByChiNhanh(optionValue);
        });

        $('#btnSave').off('click').on('click', function () {
            if ($('#frmSaveData').valid()) {
                userController.saveData();
            }
        });

        $('#btnSearch').off('click').on('click', function () {
            userController.LoadData(true);
        });

        $('#txtNameS').off('keypress').on('keypress', function (e) {
            if (e.which == 13)
                userController.LoadData(true);
        });

        $('#btnReset').off('click').on('click', function () {
            $('#txtNameS').val('');
            $('#ddlStatusS').val('');
            userController.LoadData(true);
        });

        $('.btn-edit').off('click').on('click', function () {
            $('#modalAddUpdate').modal('show');
            var id = $(this).data('id');
            // $('#txtHidID').val(id);

            $('#txtUsername').prop("disabled", true);
            userController.loadDdlChiNhanh();
            userController.loadDdlDaiLy();
            userController.loadDetail(id);



            //if ($('#ddlNhom').val() != "Users") {

            //    $('#ddlChiNhanh').prop("disabled", true);
            //    $('#ddlDaiLy').prop("disabled", true);
            //}

        });

        $('.btn-delete').off('click').on('click', function () {
            var id = $(this).data('id');
            bootbox.confirm({
                title: "Delete Confirm?",
                message: "Bạn có muốn xóa User này không?",
                buttons: {
                    cancel: {
                        label: '<i class="fa fa-times"></i> Cancel'
                    },
                    confirm: {
                        label: '<i class="fa fa-check"></i> Confirm'
                    }

                },
                callback: function (result) {
                    if (result) {
                        userController.deleteUser(id);
                    }
                }

            })
        });

        //$("#txtNgaySinh, #txtNgayCMND, #txtHanTheHDV, #txtHieuLucHoChieu, #txtHanVisa").datepicker({
        //    changeMonth: true,
        //    changeYear: true,
        //    dateFormat: "dd/mm/yy"

        //});

    },

    loadDdlChiNhanh: function () {
        $('#ddlChiNhanh').html('');
        var option = '';
        // option = option + '<option value=select>Select</option>';

        $.ajax({
            url: '/account/GetAllChiNhanh',
            type: 'GET',
            dataType: 'json',
            success: function (response) {
                //console.log(response.data);
                //if (response.length > 0) {
                //var data = JSON.stringify(response.data);
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
        //homeController.resetForm();
    },

    loadDdlDaiLy: function () {
        $('#ddlDaiLy').html('');
        var option = '';
        // option = option + '<option value=select>Select</option>';

        $.ajax({
            url: '/account/GetAllDmDaiLy',
            type: 'GET',
            dataType: 'json',
            success: function (response) {
                //if (response.length > 0) {
                //var data = JSON.stringify(response.data);
                var data = JSON.parse(response.data);
                //console.log(data);
                //$('#ddlDaiLy').html('');
                //var options = '';
                //options += '<option value="Select">Select</option>';

                //for (var i = 0; i < data.length; i++) {
                //    options += '<option value="' + data[i].SoXe + '">' + data[i].LoaiXe + '</option>'; 

                //}
                //$('#ddlDMXe').append(options);

                $.each(data, function (i, item) {
                    option = option + '<option value="' + item.Daily + '">' + item.Daily + '</option>'; //chinhanh1

                });
                $('#ddlDaiLy').html(option);
                //suachuaController.registerEvent();
                //$('#ddlDMXe').change(function () {
                //    var a=$('#ddlDMXe').val();
                //    alert(a);
                //})
                //}
            }
        });
        //homeController.resetForm();
        //var id = $('.btn-edit').data('id');

    },

    loadDdlDaiLyByChiNhanh: function (optionValue) {
        $('#ddlDaiLy').html('');
        var option = '';

        $.ajax({
            url: '/account/GetDmdailyByChiNhanh',
            type: 'GET',
            data: {
                chinhanh: optionValue
            },
            dataType: 'json',
            success: function (response) {
                //if (response.length > 0) {
                //var data = JSON.stringify(response.data);
                var data = JSON.parse(response.data);
                //$('#ddlDaiLy').html('');

                //var options = '';
                //options += '<option value="Select">Select</option>';

                //for (var i = 0; i < data.length; i++) {
                //    options += '<option value="' + data[i].SoXe + '">' + data[i].LoaiXe + '</option>'; 

                //}
                //$('#ddlDMXe').append(options);

                $.each(data, function (i, item) {
                    option = option + '<option value="' + item.Daily + '">' + item.Daily + '</option>'; //chinhanh1

                });
                $('#ddlDaiLy').html(option);
                //suachuaController.registerEvent();
                //$('#ddlDMXe').change(function () {
                //    var a=$('#ddlDMXe').val();
                //    alert(a);
                //})
                //}
            }
        });
    },

    nextUserId: function (data) {
        $.ajax({
            url: '/User/GetNextId',
            data: {
                id: data
            },
            dataType: 'json',
            type: 'GET',
            success: function (response) {
                if (response.status) {
                    var nextId = response.data;
                    $('#txtMaHD').val(nextId);
                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Get Next User ID Infomation",
                        message: response.message
                    });
                }
            }
        });
    },
    getUserId: function () {
        $.ajax({
            url: '/User/GetLastId',
            dataType: 'json',
            type: 'GET',
            success: function (response) {
                if (response.status) {
                    //console.log(response.data);
                    var data = response.data;
                    userController.nextUserId(data);
                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Get Lats User ID Infomation",
                        message: response.message
                    });
                }
            }
        });
    },

    deleteUser: function (id) {
        $.ajax({
            url: '/account/Delete',
            data: {
                id: id
            },
            dataType: 'json',
            type: 'POST',
            success: function (response) {
                if (response.status) {
                    bootbox.alert({
                        size: "small",
                        title: "Delete Infomation",
                        message: "Đã xóa thành cong.",
                        callback: function () {
                            userController.LoadData(true);
                        }
                    })
                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Delete Infomation",
                        message: response.message
                    })
                }
            },
            error: function (err) {
                console.log(err);
            }
        });
    },

    saveData: function () {
        var username = $('#txtUsername').val();
        var password = $('#txtPassword').val();
        var hoten = $('#txtHoTen').val();
        var nhom = $('#ddlNhom').val();
        var chinhanh = $('#ddlChiNhanh').val();
        var daily = $('#ddlDaiLy').val();
        var role = $('#ddlRole').val();
        var khoi = $('#ddlKhoi').val();
        var trangthai = $('#ckTrangThai').prop('checked');

        var ngaytao = $('#txtNgayTao').val();
        if (ngaytao == "")
            ngaytao = "01/01/2018";
        var nt = $.stringToDate(ngaytao, 'dd/MM/yyyy', '/');

        var ngaycapnhat = $('#txtNgayCapNhat').val();
        var ncn = $.stringToDate(ngaycapnhat, 'dd/MM/yyyy', '/');

        var hidPass = $('#txtHidPass').val();

        //var status = $('#ckStatus').prop('checked');
        var id = $('#hidID').val();
        var user = {
            username: username,
            password: password,
            hoten: hoten,
            nhom: nhom,
            chinhanh: chinhanh,
            daily: daily,
            role: role,
            khoi: khoi,
            trangthai: trangthai,

            ngaytao: nt,
            ngaycapnhat: ncn

        }
        $.ajax({
            url: '/account/SaveData',
            data: {
                strUser: JSON.stringify(user),
                Hidid: id,
                hidPass: hidPass
            },
            dataType: 'json',
            type: 'POST',
            success: function (response) {
                if (response.status) {
                    bootbox.alert({
                        size: "small",
                        title: "Save Infomation",
                        message: "Đã Lưu Thành Công.",
                        callback: function () {
                            $('#modalAddUpdate').modal('hide');
                            userController.LoadData(true);
                        }
                    });
                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Save Infomation",
                        message: response.message
                    });

                }
            },
            error: function (err) {
                console.log(err);
            }
        });
    },

    loadDetail: function (id) {
        $.ajax({
            url: '/account/GetDetail',
            data: {
                id: id
            },
            dataType: 'json',
            type: 'GET',
            success: function (response) {
                if (response.status) {
                    var data = response.data;

                    if (data.ngaytao == null)
                        nt = "";
                    if (data.ngaytao != null) {
                        nt = $.formattedDate(new Date(parseInt(data.ngaytao.substr(6))));

                    }

                    if (data.ngaycapnhat == null)
                        ncn = "";
                    if (data.ngaycapnhat != null) {
                        ncn = $.formattedDate(new Date(parseInt(data.ngaycapnhat.substr(6))));

                    }

                    $('#hidID').val(data.username);
                    $('#txtUsername').val(data.username);
                    $('#txtPassword').val('');
                    $('#txtHoTen').val(data.hoten);
                    $('#ddlNhom').val(data.nhom);
                    $('#ddlChiNhanh').val(data.chinhanh);
                    $('#ddlDaiLy').val(data.daily);
                    $('#ddlRole').val(data.role);
                    //var gioitinh = data.phai;
                    console.log(data.chinhanh);
                    console.log(data.daily);
                    //$('#ddlPhai').val(gioitinh);
                    //$('#ddlPhai').prop('selectedIndex', !gioitinh);
                    //$("#ddlPhai").selectedIndex = gioitinh;

                    $('#ddlKhoi').val(data.khoi);
                    $('#ckTrangThai').prop('checked', data.trangthai);

                    $('#txtNgayTao').val(nt);
                    $('#txtNgayCapNhat').val(ncn);
                    $('#txtHidPass').val(data.password);

                    //alert($('#ddlNhom').val());
                    if ($('#ddlNhom').val() != "Users") {

                        $('#ddlChiNhanh').prop("disabled", true);
                        $('#ddlDaiLy').prop("disabled", true);
                    }
                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Detail User Infomation",
                        message: response.message
                    });
                }
            }
        });
    },

    resetForm: function () {
        $('#hidID').val('0');
        $('#txtUsername').val('');
        $('#txtPassword').val('');
        $('#txtHoTen').val('');
        $('#ddlChiNhanh').val('');
        $('#ddlDaiLy').val('');
        $('#txtRole').val('');
        $('#txtKhoi').val('');
        $('#ckTrangThai').prop('checked', false);
        $('#txtNgayTao').val('');
        //$('#txtNgayCapNhat').val('');
        //$('#txtNgayCMND').val('');
        //$('#txtNoiCap').val('');
        //$('#txtDienThoaiDD').val('');
        //$('#txtDienThoaiNha').val('');
        //$('#txtEmail').val('');
        //$('#txtPassMail').val('');
        //$('#txtDiaChiThuongTru').val('');
        //$('#txtDiaChiTamTru').val('');
        //$('#ckHonNhan').prop('checked', false);
        //$('#txtChucDanh').val('');
        //$('#txtKinhNghiem').val('');
        //$('#txtSoTheHDV').val('');
        //$('#txtHanTheHDV').val('');
        //$('#txtHoChieu').val('');
        //$('#txtHieuLucHoChieu').val('');
        //$('#txtHanVisa').val('');
        //$('#txtGhiChuVisa').val('');
        //$('#txtNgoaiNgu').val('');
        //$('#txtChiNhanh').val('');
        //$('#txtTrinhDo').val('');
        //$('#txtTruong1').val('');
        //$('#txtHe1').val('');
        //$('#txtNganh1').val('');
        //$('#txtNam1').val('');
        //$('#txtTruong2').val('');
        //$('#txtHe2').val('');
        //$('#txtNganh2').val('');
        //$('#txtNam2').val('');
        //$('#txtTenThanNhan').val('');
        //$('#txtQuanHe').val('');
        //$('#txtDTQuanHe').val('');
        //$('#txtLyLich').val('');
    },


    LoadData: function (changePageSize) {
        var name = $('#txtNameS').val();
        var status = $('#ddlStatusS').val();

        $.ajax({
            url: '/account/LoadData',
            type: 'GET',
            data: {
                name: name,
                status: status,
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
                        //usage:
                        //var formattedDate = $.formattedDate(new Date(parseInt(item.ngaysinh.substr(6))));
                        //alert(formattedDate)
                        var ns = "";
                        if (item.ngaysinh == null)
                            ns = "";
                        else
                            ns = $.formattedDate(new Date(parseInt(item.ngaysinh.substr(6))));

                        html += Mustache.render(template, {
                            username: item.username,
                            hoten: item.hoten,
                            nhom: item.nhom,
                            daily: item.daily,
                            chinhanh: item.chinhanh,
                            trangthai: item.trangthai == true ? "<span class=\"label label-success\">Kích hoạt</span>" : "<span class=\"label label-danger\">Khóa</span>"
                        });

                    })

                    $('#tblData').html(html);
                    userController.paging(response.total, function () {
                        userController.LoadData();
                    }, changePageSize);
                    userController.registerEvent();
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
userController.init();