/////////////////////////////////////////////////////////////////////////////////////////////////
//                                   Fonctions utiles
/////////////////////////////////////////////////////////////////////////////////////////////////

function deplacement(ID_destination,ID_fichier) {
 
  // Récupère le fichier
  var fichier = DriveApp.getFileById(ID_fichier);

  // Récupère le dossier de destination
  var destinationFolder = DriveApp.getFolderById(ID_destination);
  
  // Ajoute le fichier au dossier de destination
  destinationFolder.addFile(fichier);
  
  // Supprime le fichier de tous ses dossiers parents d'origine
  var parents = fichier.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(fichier);
  }
}

function ligne_BDD(nom_BDD, colonne_recherchee, element_recherche){
  // Renvoie  si rien n'a été trouvé

  feuille = SpreadsheetApp.getActive().getSheetByName(nom_BDD);
  indice = 0;
  last = feuille.getLastRow();
  for(let i=3;i<=last;i++){
    if(feuille.getRange(i,colonne_recherchee).getValue() == element_recherche){
      indice = i;
      break
    }
  }
  return indice;
}

function id_avec_nom_prenom(nom,prenom){
  feuille = SpreadsheetApp.getActive().getSheetByName("BDD_RH");
  pos_Nom = ligne_BDD("BDD_RH",2,nom.toUpperCase());
  pos_Prenom = ligne_BDD("BDD_RH",3,prenom);
  if(pos_Nom == pos_Prenom && pos_Nom != 0){
    return feuille.getRange(pos_Nom,1).getValue();
  }
  else{
    return 0
  }
}

function id_service(nom){
  feuille = SpreadsheetApp.getActive().getSheetByName("BDD_Services");
  pos = ligne_BDD("BDD_Services",2,nom);
  return feuille.getRange(pos,1).getValue();
}

function id_equipe(nom){
  feuille = SpreadsheetApp.getActive().getSheetByName("BDD_Equipes");
  pos = ligne_BDD("BDD_Equipes",2,nom);
  return feuille.getRange(pos,1).getValue();
}

function id_poste(nom){
  feuille = SpreadsheetApp.getActive().getSheetByName("BDD_Postes");
  pos = ligne_BDD("BDD_Postes",4,nom);
  return feuille.getRange(pos,1).getValue();
}

function max_colonne(colonne,nom_feuille){
  feuille=SpreadsheetApp.getActive().getSheetByName(nom_feuille);
  max = 0;
  last = feuille.getLastRow();
  for(let i=2;i<=last;i++){
    if(feuille.getRange(i,colonne).getValue()>max){
      max = feuille.getRange(i,colonne).getValue();
    }
  }
  return(max);
}

function extraire_Id_doc(url){
  // Source : ChatGPT. Fonctionne bien
  let model = /\/d\/([a-zA-Z0-9-_]+)\//;
  let correspondance = url.match(model);
  return correspondance ? correspondance[1] : null;
}