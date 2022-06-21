$(document).ready(function(){
    console.log(1);
    console.log($('.big-card').length);
    $('.big-card').click(function(){
        console.log(2);
        $(this).toggleClass('flip');
    });
});