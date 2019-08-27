$(".InputItems").bind({
    keydown: function(e) {
        if (e.which === 45 || e.which ===189) {
            return false;
        }
        else if (e.which === 38 || e.which === 40) {
        e.preventDefault();
        }
        else if (e.which===48 && $(this).val()==0)
        {
            return false;
        }
        return true;
    }
});
