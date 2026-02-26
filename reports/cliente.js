// var rowMenu = [
	// {
		// label:"Hide Column",
		// action:function(e, column){
			// column.hide();
		// }
	// },
	// {
		// separator:true,
	// },
	// {
		// disabled:true,
		// label:"Move Column",
		// action:function(e, column){
			// column.move("col");
		// }
	// }
// ]

// Create the header menu as a function that returns the menu array
var headerMenu = function(){
    // First, create the main menu structure
    var menu = [];
    
    // Create the submenu items for column visibility
    var columnSubmenu = [];
    var columns = this.getColumns();
    
    for(let column of columns){
        //create checkbox element using font awesome icons
        let icon = document.createElement("i");
        icon.classList.add("fas");
        icon.classList.add(column.isVisible() ? "fa-check-square" : "fa-square");
        
        //build label
        let label = document.createElement("span");
        let title = document.createElement("span");
        title.textContent = " " + column.getDefinition().title;
        label.appendChild(icon);
        label.appendChild(title);
        
        //create menu item
        columnSubmenu.push({
            label:label,
            action:function(e){
                //prevent menu closing
                e.stopPropagation();
                //toggle current column visibility
                column.toggle();
                //change menu item icon
                if(column.isVisible()){
                    icon.classList.remove("fa-square");
                    icon.classList.add("fa-check-square");
                }else{
                    icon.classList.remove("fa-check-square");
                    icon.classList.add("fa-square");
                }
            }
        });
    }
    
    // Add the main menu item with submenu
    menu.push({
        label:"Colunas",
        menu:columnSubmenu
    });
    
    return menu;
};

var table = new Tabulator("#report", {
	ajaxURL: "data.json", // !!! keep this as 'ajaxURL: "data.json",' !!!
	autoColumns:"full",
	printAsHtml:true,
    printHeader:"<h1>Example Table Header<h1>",
    printFooter:"<h2>Example Table Footer<h2>",
	autoResize:false,
	responsiveLayoutCollapseStartOpen:false,
	layout:"fitData",
	movableColumns:true,
	// rowContextMenu: rowMenu,
	initialSort:[
	],
	columnDefaults:{
		tooltip:true,
		headerMenu:headerMenu,
		headerFilter:"list",
		headerFilterParams: {
			valuesLookup:true,
			//clearable:true,
			multiselect:true,
			headerFilterLiveFilter:false
		},
		bottomCalc:"sum",
	}
});

//trigger download of data.csv file
document.getElementById("download-csv").addEventListener("click", function(){
	table.download("csv", "data.csv");
});

//trigger download of data.json file
document.getElementById("download-json").addEventListener("click", function(){
	table.download("json", "data.json");
});

//trigger download of data.xlsx file
document.getElementById("download-xlsx").addEventListener("click", function(){
	table.download("xlsx", "data.xlsx", {sheetName:"Relatorio de Clientes"});
});

//trigger download of data.pdf file
document.getElementById("download-pdf").addEventListener("click", function(){
	table.download("pdf", "data.pdf", {
		orientation:"landscape", //set page orientation to portrait
		title:"Relatorio de Clientes", //add title to report
	});
});

//trigger download of data.html file
document.getElementById("download-html").addEventListener("click", function(){
	table.download("html", "data.html", {style:true});
});

document.getElementById("print-table").addEventListener("click", function(){
   table.print(false, true);
});