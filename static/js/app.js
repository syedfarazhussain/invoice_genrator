$('document').ready(function(){
 

    let leftSidebar = $('.left-sidebar').find('li');
    let title = $("#topTitle");
    leftSidebar.each((index, item)=>{
        
        let element = item.children[0];
        let trimedText = element.innerText.toLowerCase().replaceAll("\\s+", " ")

        console.log("element" + trimedText.trim());
        if(trimedText.trim() == title[0].innerText.toLowerCase()) {
            element.className = 'nav-link active';
            console.log("activeeee" + element.innerText);

        } else {

            element.className = 'nav-link link-dark';

        }
    });
    
});