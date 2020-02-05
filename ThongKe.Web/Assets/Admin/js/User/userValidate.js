///////jquery validation form
$('#frmSaveData').validate({
    errorClass: 'help-block animation-slideDown', // You can change the animation class for a different entrance animation - check animations page
    errorElement: 'div',
    errorPlacement: function (error, e) {
        e.parents('.form-group > div').append(error);
    },

    success: function (e) {
        e.closest('.form-group').removeClass('has-success has-error');
        e.closest('.help-block').remove();
    },
    rules: {
        UserName: {
            required: true,
            minlength: 5
        },
        Password: {
            required: true,
            min: 6
        }
    },
    messages: {
        UserName: {
            required: "Bạn phải nhập Username",
            minlength: "Tên phải có ít nhất 5 ký tự"
        },
        Password: {
            required: "Bạn phải nhập Password",
            min: "Password của bạn phải có ít nhất 6 ký tự"
        }
    }
});
