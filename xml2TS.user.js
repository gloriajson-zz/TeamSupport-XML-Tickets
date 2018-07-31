// ==UserScript==
// @name         ExcelToTS
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Creates multiple new tickets with data from an excel file
// @author       Gloria
// @grant        none
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Dashboard*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/TicketTabs*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Tasks*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/KnowledgeBase*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Wiki*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Search*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/WaterCooler*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Calendar*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Users*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Groups*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Customer*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Product*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Asset*
// @exclude      https://app.teamsupport.com/vcr/*/Pages/Report*
// @exclude      https://app.teamsupport.com/vcr/*/TicketPreview*
// @exclude      https://app.teamsupport.com/vcr/*/Images*
// @exclude      https://app.teamsupport.com/vcr/*/images*
// @exclude      https://app.teamsupport.com/vcr/*/Audio*
// @exclude      https://app.teamsupport.com/vcr/*/Css*
// @exclude      https://app.teamsupport.com/vcr/*/Js*
// @exclude      https://app.teamsupport.com/Services*
// @exclude      https://app.teamsupport.com/frontend*
// @exclude      https://app.teamsupport.com/Frames*
// @match        https://app.teamsupport.com/vcr/*
// @require      //maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css
// @require      https://cdn.jsdelivr.net/bootstrap.native/2.0.1/bootstrap-native.js

// ==/UserScript==

// constants
var url = "https://app.teamsupport.com/api/xml/";
var orgID = "";
var token = "";
var xhr = new XMLHttpRequest();
var parser = new DOMParser();

document.addEventListener('DOMContentLoaded', main(), false);

function main(){
    console.log("main");
  if(document.getElementsByClassName('btn-toolbar').length == 1){
    var toolbar = document.getElementsByClassName("btn-toolbar")[0]
    var button = document.createElement("button");
    button.appendChild(document.createTextNode("Mass Tickets"));
    button.setAttribute("class", "btn btn-primary");
    button.setAttribute("href", "#");
    button.setAttribute("type", "button");
    button.setAttribute("data-toggle", "modal");
    button.setAttribute("data-target", "#excelTickets");
    //button.setAttribute("data-backdrop", "static");
    toolbar.appendChild(button);
  }


  createModal();

  button.addEventListener('click', function(e){
    console.log("trying to clear text");
    e.preventDefault();
    var clear = document.getElementById('file-text').value;
    console.log(clear);
    if(clear){
        console.log("resetting text box value");
        document.getElementById('file-text').setAttribute("value", "");
        console.log(document.getElementById('file-text').value);
    }
  });

  // if Save was clicked then send a post request
  document.getElementById('create-btn').onclick = function create() {
    var customer = document.getElementById('form-select-customer').value;
    var product = document.getElementById('form-select-product').value;
    var tickets = readExcel(customer, product);
  }
}

function createModal(){
    console.log("createModal");
    // create Resolved Versions modal pop up
    var modal = document.createElement("div");
    modal.className = "modal fade";
    modal.setAttribute("id", "excelTickets");
    modal.role = "dialog";
    modal.setAttribute("tabindex", -1);
    modal.setAttribute("aria-labelledby", "excelTickets");
    modal.setAttribute("aria-hidden", true);
    document.body.appendChild(modal);

    var modalDialog = document.createElement("div");
    modalDialog.className = "modal-dialog";
    modalDialog.setAttribute("role","document");
    modal.appendChild(modalDialog);

    var modalContent = document.createElement("div");
    modalContent.className = "modal-content";
    modalDialog.appendChild(modalContent);

    //create modal header
    var modalHeader = document.createElement("div");
    modalHeader.className = "modal-header";
    modalContent.appendChild(modalHeader);

    // create header title
    var header = document.createElement("h4");
    header.className = "modal-title";
    var hText = document.createTextNode("Create New Tickets");
    header.appendChild(hText);
    modalHeader.appendChild(header);

    // create header close button
    var hbutton = document.createElement("button");
    hbutton.setAttribute("type", "button");
    hbutton.className = "close";
    hbutton.setAttribute("data-dismiss", "modal");
    hbutton.setAttribute("aria-label", "Close");
    var span = document.createElement("span");
    span.setAttribute("aria-hidden", true);
    span.innerHTML = "&times;";
    hbutton.appendChild(span);
    header.appendChild(hbutton);

    // create form in modal body
    var modalBody = document.createElement("form");
    modalBody.className="modal-body";
    modalBody.id = "create-body";
    modalContent.appendChild(modalBody);

    populateForm();

    //create modal footer
    var modalFooter = document.createElement("div");
    modalFooter.className = "modal-footer";
    modalContent.appendChild(modalFooter);

    // create save and close buttons in modal footer
    var xbtn = document.createElement("button");
    var create = document.createTextNode("Create Tickets");
    xbtn.appendChild(create);
    xbtn.id = "create-btn";
    xbtn.type = "button";
    xbtn.setAttribute("data-dismiss", "modal");
    xbtn.className = "btn btn-primary";
    var cbtn = document.createElement("button");
    var close = document.createTextNode("Close");
    cbtn.appendChild(close);
    cbtn.type = "button";
    cbtn.className = "btn btn-default";
    cbtn.setAttribute("data-dismiss", "modal");
    modalFooter.appendChild(xbtn);
    modalFooter.appendChild(cbtn);
}

function populateForm(){
  console.log("populate form");
  //create customer dropdown with options from API
  var modalBody = document.getElementById("create-body");
  var cdropdown = document.createElement("div");
  cdropdown.className = "form-group";
  var clabel = document.createElement("label");
  clabel.setAttribute("for","form-select-customer");
  clabel.innerHTML = "Select a Customer";
  var cselect = document.createElement("select");
  cselect.className = "form-control";
  cselect.setAttribute("id", "form-select-customer");

  cdropdown.appendChild(clabel);
  cdropdown.appendChild(cselect);
  modalBody.appendChild(cdropdown);

  var customers = getCustomers();
  for(var n=0; n<customers.name.length; ++n){
    var option = document.createElement("option");
    option.setAttribute("value", customers.id[n].innerHTML);
    option.innerHTML = customers.name[n].innerHTML;
    cselect.appendChild(option);
  }

  //create product dropdown with options from API
  var pdropdown = document.createElement("div");
  pdropdown.className = "form-group";
  var plabel = document.createElement("label");
  plabel.setAttribute("for","form-select-product");
  plabel.innerHTML = "Select a Product";
  var pselect = document.createElement("select");
  pselect.className = "form-control";
  pselect.setAttribute("id", "form-select-product");

  pdropdown.appendChild(plabel);
  pdropdown.appendChild(pselect);
  modalBody.appendChild(pdropdown);

  var products = getProducts();
  for(var i=0; i<products.name.length; ++i){
    var poption = document.createElement("option");
    poption.setAttribute("value", products.id[i].innerHTML);
    poption.innerHTML = products.name[i].innerHTML;
    pselect.appendChild(poption);
  }

  // create file upload button and file text bar
  var fbutton = document.createElement("button");
  fbutton.setAttribute("onClick", "document.getElementById('file-input').click();");
  fbutton.setAttribute("type", "button");
  var create = document.createTextNode("Choose XML File");
  fbutton.appendChild(create);
  var finput = document.createElement("input");
  finput.setAttribute("id", "file-input");
  finput.setAttribute("type", "file");
  finput.setAttribute("name", "file");
  finput.setAttribute("style", "display: none;");
  finput.setAttribute("accept", "text/xml");
    //.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel
  finput.onchange = function(){
      //grab the name of the file
      console.log("grabbing file result");
      console.log(document.getElementById('file-input').files[0].name);
      var name = document.getElementById('file-input').files[0].name;
      document.getElementById('file-text').setAttribute('value', name);
  };

  var ftext = document.createElement("input");
  ftext.setAttribute("type", "text");
  ftext.setAttribute("id", "file-text");
  ftext.setAttribute("placeholder", "No file selected");
  ftext.setAttribute("readonly", "true");

  modalBody.appendChild(finput);
  modalBody.appendChild(fbutton);
  modalBody.appendChild(ftext);

  console.log(modalBody);
}

function getProducts(){
  //get all the products thorugh the API
  var queryURL = url + "Products";
  xhr.open("GET", queryURL, false, orgID, token);
  xhr.send();
  var xmlDoc = parser.parseFromString(xhr.responseText,"text/xml");
  var productID = xmlDoc.getElementsByTagName("ProductID");
  var productName = xmlDoc.getElementsByTagName("Name");

  return {
    id: productID,
    name: productName
  };
}

function getCustomers(){
  //get all the customers through the API
  console.log("get customers");
  var queryURL = url + "Customers";
  xhr.open("GET", queryURL, false, orgID, token);
  xhr.send();
  var xmlDoc = parser.parseFromString(xhr.responseText,"text/xml");
  var customerID = xmlDoc.getElementsByTagName("OrganizationID");
  var customerName = xmlDoc.getElementsByTagName("Name");

  return {
    id: customerID,
    name: customerName
  };
}

function readExcel(customer, product){
  // parse through the chosen excel file
    console.log("read excel");
    var reader = new FileReader();
    var file = document.getElementById('file-input').files[0];
    reader.readAsText(file);

    console.log("before excel listener");
    reader.onloadend = function (e) {
      var xmlDoc = e.target.result;
      var xml = parser.parseFromString(xmlDoc,"text/xml");
      console.log("LOADING");
      createTickets(xml, customer, product);
    };
}

function createTickets(tickets, customer, product){
    console.log("create tickets");
    // loop through the tickets array and update their versions
    console.log(tickets);
    var len = tickets.getElementsByTagName("ticket").length;
    console.log(len);
    for(var t=0; t<len; ++t){
        var ticket = tickets.getElementsByTagName("ticket")[t];
        var title = ticket.getElementsByTagName("title")[0].innerHTML;
        var est = Number(ticket.getElementsByTagName("estimatedDays")[0].innerHTML).toFixed(2);
        var priority = ticket.getElementsByTagName("priority")[0].innerHTML;
        var id = ticket.getElementsByTagName("id")[0].innerHTML;

        if(id != null || id != undefined){
          title = title + " ("+ id + ")";
        }

        if(priority == '0' || priority == '1'){
          priority = "High";
        }else if(priority == '2'){
          priority = "Medium";
        }else{
          priority = "Low";
        }

        var data =
          '<Ticket>' +
            '<TicketStatusID>55067</TicketStatusID>' +
            '<CustomerID>' + customer + '</CustomerID>'+
            '<ProductID>' + product + '</ProductID>'+
            '<Name>' + title + '</Name>'+
            '<Estimatedevdays>' + est + '</Estimatedevdays>'+
            '<Severity>' + priority + '</Severity>'+
          '</Ticket>';

        var xmlData = parser.parseFromString(data,"text/xml");
        console.log(xmlData);
        var putURL = url + "tickets";
        xhr.open("POST", putURL, false, orgID, token);
        xhr.send(xmlData);
    }

    //force reload so website reflects resolved version change
    location.reload();
}
