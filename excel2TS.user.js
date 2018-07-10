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

function createModal(){
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
    modalHeader.appendChild(hbutton);

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

function main(){
  if(document.getElementsByClassName('btn-toolbar').length == 1){
    var toolbar = document.getElementsByClassName("btn-toolbar")[0]
    var button = document.createElement("button");
    button.appendChild(document.createTextNode("Excel Tickets"));
    button.setAttribute("class", "btn btn-primary");
    button.setAttribute("href", "#");
    button.setAttribute("data-toggle", "modal");
    button.setAttribute("data-target", "#excelTickets");
    toolbar.appendChild(button);
  }

  createModal();

  button.addEventListener('click', function(e){
    e.preventDefault();
    //var sel = document.getElementById('create-body');
    //if(sel) sel.innerHTML = "";
    //getOptions();
  });
}

function populateForm(){
  //create customer dropdown
  var modalBody = document.getElementById("create-body");
  var cdropdown = document.createElement("div");
  cdropdown.className = "form-group";
  var clabel = document.createElement("label");
  clabel.setAttribute("for","form-select");
  clabel.innerHTML = "Select a Customer";
  var cselect = document.createElement("select");
  cselect.className = "form-control";
  cselect.setAttribute("id", "form-select");

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

  //create product dropdown
  var pdropdown = document.createElement("div");
  pdropdown.className = "form-group";
  var plabel = document.createElement("label");
  plabel.setAttribute("for","form-select");
  plabel.innerHTML = "Select a Product";
  var pselect = document.createElement("select");
  pselect.className = "form-control";
  pselect.setAttribute("id", "form-select");

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

  //create file selection
  /*var file = document.createElement("div");
  file.className = "form-group";
  var flabel = document.createElement("label");
  flabel.setAttribute("for","form-control");
  flabel.innerHTML = "Select an Excel file";
  var fselect = document.createElement("input");
  fselect.setAtribute("type", "file");
  fselect.setAttribute("class", "form-control-file");
  fselect.setAttribute("id", "form-control");

  file.appendChild(flabel);
  file.appendChild(fselect);
  modalBody.appendChild(file);*/
}

function getProducts(){
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
