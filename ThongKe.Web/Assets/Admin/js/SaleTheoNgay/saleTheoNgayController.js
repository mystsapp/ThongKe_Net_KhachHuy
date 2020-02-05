var saleTheoNgayController = {
    init: function () {
        // saleTheoNgayController.LoadData();t
        saleTheoNgayController.registerEvent();
    },

    registerEvent: function () {

        $('#frmSearch').validate({
            rules: {
                tungay: {
                    required: true,
                    date: true
                },
                denngay: {
                    required: true,
                    date: true
                }
            },
            messages: {
                tungay: {
                    required: "Trường này không được để trống.",
                    date: "Chưa đúng định dạng."
                },
                denngay: {
                    required: "Trường này không được để trống.",
                    //number: "password phải là số",
                    date: "Chưa đúng định dạng."
                }
            }
        });


        $('#btnSearch').off('click').on('click', function () {
            $('#frmSearch').submit();
        });

        $('#btnReset').off('click').on('click', function () {
            saleTheoNgayController.resetForm();
        });

        

        $("#txtTuNgay, #txtDenNgay").datepicker({
            changeMonth: true,
            changeYear: true,
            dateFormat: "mm/dd/yy"

        });
    },
    resetForm: function () {
        $('#txtTuNgay').val('');
        $('#txtDenNgay').val('');
    },

    loadDdlDaiLyByChiNhanh: function (optionValue) {
        $('#ddlDaily').html('');
        var option = '';
        option = '<option value=" ">' + "Tất cả" + '</option>';
        $.ajax({
            url: '/baocao/GetDmdailyByChiNhanh',
            type: 'GET',
            data: {
                chinhanh: optionValue
            },
            dataType: 'json',
            success: function (response) {
              
                var data = JSON.parse(response.data);
             
                $.each(data, function (i, item) {
                    option = option + '<option value="' + item.Daily + '">' + item.Daily + '</option>'; //chinhanh1

                });
                $('#ddlDaiLy').html(option);
            
            }
        });
    },

}
saleTheoNgayController.init();