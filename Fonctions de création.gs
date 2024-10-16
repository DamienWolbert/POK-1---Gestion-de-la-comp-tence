/////////////////////////////////////////////////////////////////////////////////////////////////
//                                   Fonctions de création
/////////////////////////////////////////////////////////////////////////////////////////////////

function Creation_poste() {
  // Vérification des données d'entrée
  let feuille = SpreadsheetApp.getActive().getSheetByName("New");
  let service = feuille.getRange(4,2).getValue();
  let equipe = feuille.getRange(5,2).getValue();
  let nom = feuille.getRange(3,2).getValue();
  let poste_hors_service = feuille.getRange(6,2).getValue();

  // Entrées

  if(poste_hors_service == ""){
    SpreadsheetApp.getUi().alert("Veuillez indiquer si le poste sélectionné fait partie d'un service");
    return;
  }

  if(nom == "") {
    SpreadsheetApp.getUi().alert("Veuillez entrer un nom de poste.");
    return;
  }

  if(service == "" && equipe != ""){
    SpreadsheetApp.getUi().alert("Vous avez indiqué une équipe sans y indiquer de service.");
    return;   
  }

  // Existance poste
  if(ligne_BDD("BDD_Postes", 4, nom)!=0){
    SpreadsheetApp.getUi().alert("Le poste que vous tentez de créer existe déjà.");
    return;
  }

  // Complétion de la BDD
  let bdd = SpreadsheetApp.getActive().getSheetByName("BDD_Postes");
  let last = bdd.getLastRow();
  bdd.getRange(last+1,1).setValue(max_colonne(1,"BDD_Postes")+1);
  bdd.getRange(last+1,4).setValue(nom);

  //if(poste_hors_service != "Oui"){// Cas avec service et équipe
    if(service != ""){
      let id_serv = id_service(service);
      bdd.getRange(last+1,2).setValue(id_serv);
    }
    if(equipe != ""){
      let id_eq = id_equipe(equipe);
      bdd.getRange(last+1,3).setValue(id_eq);
    }
  
  //else{

}

function Creation_profile(){
  let feuille = SpreadsheetApp.getActive().getSheetByName("New");
  let nom = feuille.getRange(9,2).getValue();
  let prenom = feuille.getRange(10,2).getValue();
  let service = feuille.getRange(11,2).getValue();
  let equipe = feuille.getRange(12,2).getValue();
  let poste = feuille.getRange(13,2).getValue();
  let admin = feuille.getRange(14,2).getValue();

  // Vérification des données d'entrée
  if(nom=="" || prenom==""){
    SpreadsheetApp.getUi().alert("Veuillez entrer un nom et un prenom");
    return;
  }

  if(admin == ""){
    SpreadsheetApp.getUi().alert("Veuillez préciser si le profile est un profile administrateur");
    return;
  }

  // Profil existant
  if(ligne_BDD("BDD_RH", 2, nom)==ligne_BDD("BDD_RH", 3, prenom) && ligne_BDD("BDD_RH", 2, nom)!=0){
    SpreadsheetApp.getUi().alert("Le profile que vous tentez de créer existe déjà.");
    return;
  }

  //Complétion de la BDD
  let BDD = SpreadsheetApp.getActive().getSheetByName("BDD_RH");
  let last = BDD.getLastRow();
  let id_nouveau = max_colonne(1,"BDD_RH") + 1;
  BDD.getRange(last+1,1).setValue(id_nouveau);
  BDD.getRange(last+1,2).setValue(nom);
  BDD.getRange(last+1,3).setValue(prenom);

  if(poste!=""){
    BDD.getRange(last+1,6).setValue(id_poste(poste));
  }

  if(service!=""){
    BDD.getRange(last+1,4).setValue(id_service(service));
  }

  if(equipe!=""){
    BDD.getRange(last+1,5).setValue(id_equipe(equipe));
  }

  if(admin=="Oui"){
    let BDD_admin = SpreadsheetApp.getActive().getSheetByName("BDD_Admin");
    let last_admin = BDD_admin.getLastRow();
    BDD_admin.getRange(last_admin+1,1).setValue(max_colonne(1,"BDD_Admin"));
    BDD_admin.getRange(last_admin+1,2).setValue(id_nouveau);
  }

  let url_indiv = Creation_fiche_individuelle(nom, prenom, service, equipe, poste);
  BDD.getRange(last+1,7).setValue(url_indiv);
}

function Creation_equipe(){
  feuille = SpreadsheetApp.getActive().getSheetByName("New");
  nom = feuille.getRange(22,2).getValue();
  nom_respo = feuille.getRange(24,2).getValue();
  prenom_respo = feuille.getRange(25,2).getValue();
  nom_service = feuille.getRange(23,2).getValue();

  // Vérification des données d'entrée
  
  // Entrée du nom
  if(nom == ""){
    SpreadsheetApp.getUi().alert("Veuillez entrer un nom d'équipe.");
    return;
  }

  // Entrée service
  if(nom_service == ""){
    SpreadsheetApp.getUi().alert("Veuillez associer cette équipe à un service");
    return;
  }

  // Equipe existante
  if(ligne_BDD("BDD_Equipes", 2, nom)!=0){
    SpreadsheetApp.getUi().alert("L'équipe que vous essayez de créer existe déjà.");
    return;
  }

  // Complétion BDD
  let bDD=SpreadsheetApp.getActive().getSheetByName("BDD_Equipes");
  let num = max_colonne(1,"BDD_Equipes")+1;
  let last = bDD.getLastRow();
  bDD.getRange(last+1,1).setValue(num);
  bDD.getRange(last+1,2).setValue(nom);
  bDD.getRange(last+1,3).setValue(id_service(nom_service));
  
  // Responsable
  let id_respo = id_avec_nom_prenom(nom_respo, prenom_respo);
  if(id_respo!=0){
    bDD.getRange(last+1,4).setValue(id_respo);
  }
}

function Creation_service(){
  let feuille = SpreadsheetApp.getActive().getSheetByName("New");
  let nom = feuille.getRange(17,2).getValue();
  let nom_respo = feuille.getRange(18,2).getValue();
  let prenom_respo = feuille.getRange(19,2).getValue();
  let nom_service = feuille.getRange(23,2).getValue();

  // Vérification des données d'entrée

  // Entrée du nom
  if(nom == ""){
    SpreadsheetApp.getUi().alert("Veuillez entrer un nom de service.");
    return;
  }

  // Service existant
  if(ligne_BDD("BDD_Services", 2, nom)!=0){
    SpreadsheetApp.getUi().alert("Le service que vous essayez de créer existe déjà.");
    return;
  }

  // Complétion de la BDD
  BDD = SpreadsheetApp.getActive().getSheetByName("BDD_Services");
  let last = BDD.getLastRow();
  let num = max_colonne(1,"BDD_Services")+1;
  BDD.getRange(last+1,1).setValue(num);
  BDD.getRange(last+1,2).setValue(nom);

  //Responsable
  let id_respo = id_avec_nom_prenom(nom_respo, prenom_respo);
  if(id_respo != 0){
    BDD.getRange(last+1,3).setValue(id_respo);
  }
}

function Creation_fiche_individuelle(nom, prenom, service, equipe, poste){  
  // Copie
  let tem_url = "https://docs.google.com/spreadsheets/d/1jz4QhZDmInLyTs9LAemzIZA6sGFmrF87wcNzQJi8ugY/edit?gid=1839674467#gid=1839674467";
  let tem_id = extraire_Id_doc(tem_url);
  let tem = DriveApp.getFileById(tem_id);
  let dossier_Id = "1jLOU9tI0f6tfL37XP5qupG-zpnBxYIZs";
  let dossier = DriveApp.getFolderById(dossier_Id);

  let nouveau = tem.makeCopy(nom + "_" + prenom,dossier);
  let url = nouveau.getUrl();

  // Page de garde
  let feuille = SpreadsheetApp.openByUrl(url).getSheetByName("Page de garde");
  feuille.getRange(10,4).setValue(nom);
  feuille.getRange(11,4).setValue(prenom);
  feuille.getRange(12,4).setValue(service);
  feuille.getRange(13,4).setValue(equipe);
  feuille.getRange(14,4).setValue(poste);

  return url;
}