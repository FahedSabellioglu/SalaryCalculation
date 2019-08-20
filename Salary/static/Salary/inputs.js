function InputControl(id){
        selectedElement = $('#'+id);
        $('td input').each(function(){
            if($(this).attr('id') > id ){
                $(this).val(selectedElement.val())
            }
        })

}