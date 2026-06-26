//#include './extendables/extendables.jsx'


var doc;

var artistes = [
    ["BARRY", 19],
    ["GARCIA", 21, 20],
    ["LAKHRISSI", 22],
    ["BAHRI", 23],
    ["ONO", 24, 40],
    ["HATOUM", 25],
    ["MOREL", 26, 20],
    ["RUTAULT", 27],
    ["LEWITT", 28],
    ["CLOSKY", 29, 50],
    ["HIGGINS", 31, 30],
    ["FERRER", 33],
    ["MOLNAR", 34],
    ["LARVALABS", 35],
    ["BRECHT", 36],
    ["PAIK", 37, 20],
    ["PAPANEK", 38, 20],
    ["BERNIER", 39, 50],
    ["WURM", 40],
    ["HYBER", 41, 20],
    ["FRIEDMAN", 42],
    ["AYCOCK", 43, 30],
    ["BLAZY", 44],
    ["PHINTONG", 45, 20],
    ["FONTAINE", 46],
    ["FRAMIS", 47, 20],
    ["KIWANGA", 48, 20],
    ["MISPLEAERE", 49],
    ["HOLLER", 50],
    ["ECHAKHCH", 51], 
    ["SALADIN", 53, 50],
    ["ATTIA", 55, 0],
    ["VANDEBROUCK", 56, 50],
    ["PIERON", 57],
    ["BUCHY", 59]
]


/*
var artistes = [
    ["BARRY", 19],
    ["GARCIA", 21, 20],
    ["LAKHRISSI", 22],
    ["BAHRI", 23],
    ["ONO", 24, 40]
]
*/

var fichiers_artistes = [];

var chemin_maquette = '/Users/kevindonnot/Desktop/__MAMCS_COMPILER_INTERIEUR/240904_MAMCS_INTERIEUR_EXE.indd';
var chemin_dessins = '/Users/kevindonnot/Desktop/240904_DESSINS_DEF/';
var chemin_export = '/Users/kevindonnot/Desktop/__MAMCS_COMPILER_INTERIEUR/GENERATIONS/';

var ex = 0;
var ex_debut = 780;
var n_exemplaires = 20;



//////////////////////////////////////////
for (var i=ex_debut; i<ex_debut+n_exemplaires; i++) {
    app.open(chemin_maquette)
    doc = app.activeDocument;
    ex = i
    
    if (i == ex_debut) {
        charger_images();
    }
    
    poser_images();
    //poser_image_couverture();
    multiplier_farkas();
    changer_calque_actif();

    poser_numerotation();
    poser_numerotation_couverture();
    exporter_pdf();

    doc.close(SaveOptions.NO)
}


//////////////////////////////////////////
/*
doc = app.activeDocument;
charger_images();
poser_images();
multiplier_farkas();
changer_calque_actif();
 poser_numerotation();
poser_numerotation_couverture();
*/





/////////////////////////////////////////////////////////////
// CHARGER TOUS LES CHEMINS DES IMAGES
function charger_images() {
    for (var i=0; i<artistes.length; i++) {
        var artiste = artistes[i];
        var fichiers_artiste = Folder(chemin_dessins + artiste[0]+"/").getFiles("*.tif");
        fichiers_artiste.sort();
        fichiers_artistes.push(fichiers_artiste);
    }
}


/////////////////////////////////////////////////////////////
// PLACER LES IMAGES D'UN EXEMPLAIRE
function poser_images() {
    for (var i=0; i<artistes.length; i++) {
        var page = doc.pages[artistes[i][1] + 1];
        var fichier = fichiers_artistes[i][ex];

        var bloc_image = page.rectangles.add();

        var decalage_y = 0;
        if (artistes[i][2] != undefined) {
            decalage_y = artistes[i][2];
        }

        if (page.side == PageSideOptions.LEFT_HAND) {
            bloc_image.geometricBounds = [-5-decalage_y,-5,  297+5-decalage_y, 210]
        } else {
            bloc_image.geometricBounds = [-5-decalage_y,210,  297+5-decalage_y, 425]
        }
        bloc_image.fillColor = "None"
        bloc_image.place(fichier.fsName, false);
        bloc_image.fit (FitOptions.FILL_PROPORTIONALLY);
    }
}

/////////////////////////////////////////////////////////////
// PLACER L'IMAGE DE COUVERTURE
// N_EX - 1 (0 DEVIENT 800)
function poser_image_couverture() {
    var couverture = doc.pages[0];

    var artiste_en_couv = ex % artistes.length;
    var choix_image;;
    if (ex == 0) {
        choix_image = 799
    } else {
        choix_image = ex-1
    }

    var fichier = fichiers_artistes[artiste_en_couv][choix_image];
    var bloc_image = couverture.rectangles.add();

    bloc_image.geometricBounds = [-5,0,  297+5, 215]
    bloc_image.fillColor = "None"
    bloc_image.place(fichier.fsName, false);
    bloc_image.fit (FitOptions.FILL_PROPORTIONALLY);
}

/////////////////////////////////////////////////////////////
// NUMÉROTATION INTERIEURE
function poser_numerotation() {
    var gabarit = doc.masterSpreads.itemByName("X-NUMEROTATION");
    var style_courant = doc.paragraphStyles.itemByName("NUMEROTATION");

    var bloc_texte = gabarit.textFrames.add();
    
    bloc_texte.properties = {
        geometricBounds : [7.75, 160.381,  7.75+3.44, 160.381+26.179 ],
        absoluteRotationAngle: 90,
        contents : "ex. "+(ex+1)+" / 800"    
    };
    bloc_texte.parentStory.appliedParagraphStyle = style_courant;
    bloc_texte.parentStory.justification = Justification.RIGHT_ALIGN;
    bloc_texte.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN


    var bloc_texte = gabarit.textFrames.add();
    bloc_texte.properties = {
        geometricBounds : [7.75, 160.381+222,  7.75+3.44, 160.381+26.179+222 ],
        absoluteRotationAngle: 90,
        contents : "ex. "+(ex+1)+" / 800"    
    };
    bloc_texte.parentStory.appliedParagraphStyle = style_courant;
    bloc_texte.parentStory.justification = Justification.RIGHT_ALIGN;
    bloc_texte.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN
}

/////////////////
function changer_calque_actif() {
    var calque_images = doc.layers.itemByName("IMAGES");
    var calque_texte = doc.layers.itemByName("CONTENU");
    calque_images.locked = true;
    calque_texte.locked = false;
    doc.activeLayer = calque_texte;
}

/////////////////////////////////////////////////////////////
// NUMÉROTATION COUVERTURE
function poser_numerotation_couverture() {
    var page_titre = doc.pages[0];
    var style_titre = doc.paragraphStyles.itemByName("TITRE_COUV");
    var bloc_texte = page_titre.textFrames.add();
    
    var x = 185-63
    var y = 7
    var w = 63
    var h = 17
    bloc_texte.properties = {
        geometricBounds : [y, x, y+h, x+w],
        absoluteRotationAngle: 90,
        contents : (ex+1)+" / 800"    
    };
    bloc_texte.parentStory.appliedParagraphStyle = style_titre;
    bloc_texte.parentStory.justification = Justification.RIGHT_ALIGN;
    bloc_texte.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN
}

/////////////////////////////////////////////////////////////
// MULTIPLIER BLOCS FARKAS
function multiplier_farkas() {
    var page_farkas = doc.pages[55];
    var w = 80;
    var h = 20;
    bloc = page_farkas.textFrames[0];

    bloc.geometricBounds = [0, 0, h, w];
    bloc.parentStory.alignToBaseline = false;

    var x = Math.random()*(210-w);
    var y = Math.random()*(230-h);
    bloc.move([x,y]);

    n_blocs = Math.floor(Math.random()*3)
    //n_blocs += 20
    for (var i=0; i<n_blocs; i++) {
        var x = Math.random()*(210-w);
        var y = Math.random()*(230-h);
        bloc.duplicate([x,y]);
    }
}


/////////////////////////////////////////////////////////////
// EXPORTER_PDF
function exporter_pdf() {
    var numero = addLeadingZeros(ex+1, 3);
    var fichier = numero+"_INTERIEUR_MAMCS.pdf";
    var preset = app.pdfExportPresets.itemByName("PRINT_MAMCS");
    doc.exportFile(ExportFormat.PDF_TYPE, chemin_export+fichier, false, preset);  
}



function addLeadingZeros(num, targetLength) {
    var numStr = num.toString();
    while (numStr.length < targetLength) {
      numStr = '0' + numStr;
    }
    return numStr;
}





//doc.saveACopy(destination, false);
/*
var log = new Array();
doc.groups.everyItem().ungroup();  
log.push("**********");
// Pour chaque bloc lier, dupliquer pour enlever la laisison

for (a=0; a<100; a++) {

for (i=0; i<doc.stories.length; i++) {
    var blocsTexte = doc.stories.item(i).textContainers;
    if (blocsTexte.length > 1) {
        for(j = 0; j < blocsTexte.length; j++){
            blocsTexte[j].duplicate();
        }
        for(j = 0; j < blocsTexte.length; j++){
            blocsTexte[j].remove();
            //blocsTexte[j].contents.remove();
        }
    }
}

}

for (a=0; a<10; a++) {
for (i=0; i<doc.textFrames.length; i++) {
    if (doc.textFrames[i].contents != "") {
        doc.textFrames[i].createOutlines(); 
    }
    
}
}



doc.groups.everyItem().ungroup();  

for (i=0; i<doc.splineItems.length; i++) {

    var courbes = doc.splineItems.item(i);
   // log.push(courbes.fillColor.name);
    
    if (courbes.fillColor.name == "Black" || courbes.strokeColor.name == "Black") {
    
    courbes.fillColor = "None";
    courbes.strokeColor = "Black";
    courbes.strokeWeight = 0.001;
    courbes.endCap = EndCap.ROUND_END_CAP;
    courbes.endJoin = EndJoin.ROUND_END_JOIN;
    courbes.paths.everyItem().pathType = PathType.OPEN_PATH;  
    }
}

$.writeln(log.join("\n"));
*/