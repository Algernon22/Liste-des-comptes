/**
* @OnlyCurrentDoc
*/

function listeUtilisateurs() {
  var fuseauHoraire = Session.getScriptTimeZone();
  
  var classeur = SpreadsheetApp.getActive();
   
  var feuilleListe = classeur.getActiveSheet();
  feuilleListe.clearContents();

  try {
    feuilleListe.getDataRange();
    var banding = feuilleALL.getDataRange().getBandings()[0];
    banding.remove();
  } catch (e) {
    
  }
  
  var pageToken,
      page,
      count = 0;
  var listArray = [];
  listArray.push(['E-mail','Nom complet', 'Prénom', 'Nom', 'Centre de coût','Service', 'Fonction', 'E-mail Resp.', 'Type', 'UO', 'Adresse','Date Dernière connexion', 'Date de création'])
  
  
  do {
    page = AdminDirectory.Users.list({
      domain : 'faucheux.bzh',
      orderBy : 'familyName',
      projection: 'full',
      maxResults: 500,   
      pageToken : pageToken
    });
    var users = page.users;
    if (users) {
      for (var i = 0; i < users.length; i++) {
        var user = users[i];
        var service,
            centre_cout,
            type,
            adrresp,
            adresse,
            fonction,
            dateCreation;
        try {
          service = user.organizations[0].department;
        } catch (e) {
          service = e
        }
        try {
          fonction = user.organizations[0].title;
        } catch (e) {
          fonction = e
        }
        try {
          type = user.organizations[0].description;
        } catch (e) {
          type = e
        }
        try {
          branche = user.organizations[0].costCenter;
        } catch (e) {
          branche = e
        }
        try {
          localisation = user.organizations[0].location;
        } catch (e) {
          branche = e
        }        
        try {
          adrresp = user.relations[0].value;
        } catch (e) {
          adrresp = ""
        }                       
        try {
          adresse = user.addresses[0].formatted;
        } catch (e) {
          adresse = ""
        } 
        try {
          dateCreation = Utilities.formatDate(user.creationTime, fuseauHoraire, "yyyy/MM/dd'T'HH:mm:ss'Z'");
          dateCreation = dateCreation.substring(0,10); 
        } catch (e) {
          dateCreation = user.creationTime;
          dateCreation = dateCreation.substring(0,10); 
        }
        
        
        listArray.push([user.primaryEmail, user.name.fullName, user.name.givenName, user.name.familyName, centre_cout, service, fonction, adrresp, type, user.orgUnitPath, adresse, user.lastLoginTime, dateCreation]);
        
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  try {
    var feuilleListe = classeur.getActiveSheet();
    feuilleListe.getDataRange();
  } catch (err) {
    var feuilleListe = classeur.insertSheet();
  }
  
  feuilleListe.getRange(1, 1, listArray.length, listArray[0].length).setValues(listArray);
  feuilleListe.getRange(1, 6, feuilleListe.getLastRow(), 4).setHorizontalAlignment("left");
  feuilleListe.getRange(1, 1, feuilleListe.getLastRow(), 1).setHorizontalAlignment("left");
  feuilleListe.autoResizeColumns(1, feuilleListe.getLastColumn());
  feuilleListe.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  
  feuilleListe.getRange(1,1,1,feuilleListe.getLastColumn()).setFontWeight("bold");
  feuilleListe.getRange(1,1,1,feuilleListe.getLastColumn()).setBackground("#0d5973")
  feuilleListe.getRange(1,1,1,feuilleListe.getLastColumn()).setFontColor("#f086a4");
  
  // Classement
  var dernierRang = feuilleListe.getLastRow()
  var derniereColonne = feuilleListe.getLastColumn();
  var plageDonnees = feuilleListe.getRange(2, 1, dernierRang, derniereColonne);
  plageDonnees.sort({column: 1, ascending: false});
  
  var font = 'Source Sans Pro';
  plageDonnees.setFontFamily(font);
  
}
