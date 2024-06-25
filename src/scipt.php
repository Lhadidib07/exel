<?php
require 'vendor/autoload.php'; // Inclure l'autoload pour pouvoir utiliser les bibliothèques externes

use PhpOffice\PhpSpreadsheet\IOFactory; // Utiliser la classe IOFactory pour lire et écrire des fichiers Excel
use PhpOffice\PhpSpreadsheet\Spreadsheet; // Utiliser la classe Spreadsheet pour manipuler des feuilles de calcul

// Chemin du fichier source Excel (doit pointer vers un fichier spécifique, pas un dossier)
$sourceFilePath = './exelSource/monFichier.xlsx'; // Exemple de chemin vers un fichier Excel spécifique
// Chemin du dossier de destination
$dossierDestination = "exelDestination/"; // Dossier où les fichiers Excel divisés seront sauvegardés

// Fonction pour mapper les colonnes de la source vers la destination
function mapSourceColumnToDestination($sourceColumn) {
    $mapping = [
        'A' => 'X', // Mapper la colonne A vers la colonne X
        'B' => 'Y', // Mapper la colonne B vers la colonne Y
        // Ajoutez d'autres mappages selon vos besoins
    ];

    return $mapping[$sourceColumn] ?? $sourceColumn; // Retourner la colonne mappée ou la colonne source si aucun mappage trouvé
}

$spreadsheet = IOFactory::load($sourceFilePath); // Charger le fichier Excel spécifié

$worksheet = $spreadsheet->getActiveSheet(); // Obtenir la première feuille de calcul du fichier

$maxRowsPerFile = 300; // Nombre maximal de lignes par fichier Excel divisé
$rowCounter = 0; // Compteur de lignes pour le fichier actuel
$fileIndex = 1; // Index pour nommer les fichiers divisés

$newSpreadsheet = new Spreadsheet(); // Créer un nouveau fichier Excel pour la division
$newSheet = $newSpreadsheet->getActiveSheet(); // Obtenir la feuille de calcul du nouveau fichier

// Itérer sur chaque ligne de la feuille de calcul source 
// complexiter (O(n^2)) à éviter si possible 

foreach ($worksheet->getRowIterator() as $row) {
    $rowCounter++; // Incrémenter le compteur de lignes
    
    // Itérer sur chaque cellule de la ligne
    foreach ($row->getCellIterator() as $cell) {
        $destinationColumn = mapSourceColumnToDestination($cell->getColumn()); // Mapper la colonne de la cellule
        $newSheet->setCellValue($destinationColumn.$rowCounter, $cell->getValue()); // Copier la valeur dans la nouvelle feuille
    }
    
    // Vérifier si le nombre maximal de lignes par fichier est atteint
    if ($rowCounter == $maxRowsPerFile) {
        $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx'); // Préparer l'écriture du fichier Excel
        $writer->save($dossierDestination . "sous_fichier_{$fileIndex}.xlsx"); // Sauvegarder le fichier divisé
        
        $newSpreadsheet = new Spreadsheet(); // Créer un nouveau fichier Excel pour la prochaine division
        $newSheet = $newSpreadsheet->getActiveSheet(); // Obtenir la nouvelle feuille de calcul
        $rowCounter = 0; // Réinitialiser le compteur de lignes
        $fileIndex++; // Incrémenter l'index du fichier
    }
}

// Vérifier s'il reste des lignes non sauvegardées dans un fichier divisé
if ($rowCounter > 0) {
    $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx'); // Préparer l'écriture du dernier fichier Excel
    $writer->save($dossierDestination . "sous_fichier_{$fileIndex}.xlsx"); // Sauvegarder le dernier fichier divisé
}
?>