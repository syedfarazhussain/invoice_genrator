$('document').ready(function(){

    let leftSidebar = $('.left-sidebar').find('li');
    leftSidebar.each((index, item)=>{
        item.addEventListener('click', function(){
            
            if(this.find('a').hasClass('active')){
                console.log('true')
            } else {
               console.log('false')
            }
        })
        console.log("leftSidebar" + item);

    })
    
});