$("#SocialStatusSelect").on('change',function(){
    if($(this).val()==1)
    {
        $("#Married").show()
    }
    else{
        $("#Married").hide()
    }
});