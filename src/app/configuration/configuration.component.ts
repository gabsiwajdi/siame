import { Component, OnInit } from '@angular/core';
import Swal from 'sweetalert2';
import * as XLSX from 'xlsx';

//** import pdf maker */
const pdfMake = require('pdfmake/build/pdfmake');
const pdfFonts = require('pdfmake/build/vfs_fonts');
pdfMake.vfs = pdfFonts.pdfMake.vfs;

@Component({
  selector: 'app-configuration',
  templateUrl: './configuration.component.html',
  styleUrls: ['./configuration.component.scss'],
})
export class ConfigurationComponent implements OnInit {
  currentDate = '';
  employees: any[] = [];
  monthlyData: any[] = [];
  fileData: any;

  constructor() {
    const now = new Date();
    this.currentDate = now.toISOString().split('T')[0];
    this.loadEmployeesFromLocalStorage();
  }

  ngOnInit(): void {}

  saveEmployeesToLocalStorage() {
    localStorage.setItem('employees', JSON.stringify(this.employees)); // Enregistrer les employés dans le localStorage
  }

  loadEmployeesFromLocalStorage() {
    const storedEmployees = localStorage.getItem('employees');
    if (storedEmployees) {
      this.employees = JSON.parse(storedEmployees); // Charger les employés depuis le localStorage
    }
  }
  readFile(event: any): void {
    const file: File = event.target.files[0];
    const extension = file.name.split('.').pop();

    if (extension === 'csv') {
      this.readCSVFile(file);
      Swal.fire({
        icon: 'success',
        title: 'Fichier importé avec succès',
      });
    } else if (extension === 'xlsx') {
      this.readXlsxFile(file);
      Swal.fire({
        icon: 'success',
        title: 'Fichier importé avec succès',
        timer: 1500,
      });
    } else {
      Swal.fire({
        icon: 'error',
        title: 'Format de fichier non pris en charge !',
      });
      console.error('Format de fichier non pris en charge');
    }
  }

  // Lire un fichier XLSX
  readXlsxFile(file: File): void {
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: true });
      this.employees = jsonData; // Mettre à jour le tableau d'employés
      this.saveEmployeesToLocalStorage(); // Enregistrer les données mises à jour dans le local storage
      console.log(jsonData);
    };

    reader.readAsArrayBuffer(file);
  }

  readCSVFile(file: File): void {
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const contents: string = e.target.result; // Définir le type de 'contents' comme 'string'
      const lines: string[] = contents.split(/\r\n|\n/);
      const headers: string[] = lines[0]
        .split(';')
        .map((header: string) => header.trim()); // Définir le type de 'header' comme 'string'

      const jsonData: any[] = [];

      for (let i = 1; i < lines.length; i++) {
        const currentLine: string[] = lines[i]
          .split(';')
          .map((value: string) => value.trim().replace(/[\r\n]/g, '')); // Définir le type de 'value' comme 'string'

        if (currentLine.length === headers.length) {
          const obj: any = {};

          for (let j = 0; j < headers.length; j++) {
            obj[headers[j]] = this.convertValue(currentLine[j]);
          }

          jsonData.push(obj);
        }
      }
      this.employees = jsonData; // Mettre à jour le tableau d'employés
      this.saveEmployeesToLocalStorage(); // Enregistrer les données mises à jour dans le local storage
    };

    reader.readAsText(file);
  }
  convertValue(value: string): any {
    const scientificNotationRegex = /^[+\-]?\d+(\.\d+)?[eE][+\-]?\d+$/;

    if (scientificNotationRegex.test(value)) {
      return Number(value).toString();
    } else {
      return value;
    }
  }

  // Méthode pour générer le rapport PDF

  async generateRapportPDF() {
    const currentMonth = new Date().getMonth() + 1;
    const daysInMonth = new Date(
      new Date().getFullYear(),
      currentMonth,
      0
    ).getDate();
    const accessMatrix = [];
    const dateHeaders = Array.from({ length: daysInMonth }, (_, i) =>
      (i + 1).toString()
    );
    accessMatrix.push(['Employé', ...dateHeaders, 'NJ']);

    this.employees.forEach((employee: any) => {
      const employeeRow = [employee.nom];
      let totalDaysAccessed = 0;

      for (let i = 1; i <= daysInMonth; i++) {
        const currentDate = new Date(
          new Date().getFullYear(),
          currentMonth - 1,
          i
        ).toLocaleDateString();
        const storedCodes = localStorage.getItem(currentDate);
        if (storedCodes) {
          const usedCodes = JSON.parse(storedCodes);
          const hasAccess = usedCodes.includes(employee.code);
          employeeRow.push(hasAccess ? '*' : '');
          if (hasAccess) totalDaysAccessed++;
        } else {
          employeeRow.push('');
        }
      }
      employeeRow.push(totalDaysAccessed.toString());
      accessMatrix.push(employeeRow);
    });

    const currentYear = new Date().getFullYear();
    const monthNames = [
      'Janvier',
      'Février',
      'Mars',
      'Avril',
      'Mai',
      'Juin',
      'Juillet',
      'Août',
      'Septembre',
      'Octobre',
      'Novembre',
      'Décembre',
    ];
    const currentMonthName = monthNames[currentMonth - 1];
    const currentDate = new Date().toLocaleDateString('fr-FR');

    const documentDefinition = {
      content: [
        // Logo
        {
          image: await this.getBase64ImageFromURL(
            '../../assets/images/logo.png'
          ),
          width: 160, // Augmenter la taille du logo
          height: 80, // Augmenter la taille du logo
          alignment: 'left',
          margin: [0, 0, 0, 0], // Marges nulles pour éliminer l'espace
        },
        // Titre
        {
          text: 'Rapport Contine',
          style: 'header',
          alignment: 'center',
          color: 'blue', // Mettre le titre en bleu
          margin: [0, -20, 0, 10], // Ajouter des marges au titre et le déplacer vers le haut
        },
        // Mois et année
        {
          text: `Mois : ${currentMonthName} ${currentYear}`,
          alignment: 'center',

          margin: [0, 0, 0, 5], // Ajouter des marges au mois
        },
        {
          text: `Date : ${currentDate}`,
          alignment: 'center',

          margin: [0, 0, 0, 20], // Ajouter des marges à la date
        },
        // Saut de ligne
        { text: '\n\n' },
        // Tableau des données
        {
          table: {
            headerRows: 1,
            widths: Array(accessMatrix[0].length).fill('auto'),
            body: accessMatrix,
            pageBreak: 'auto',
          },
        },
      ],
      styles: {
        header: {
          fontSize: 24,
          bold: true,
          decoration: 'underline', // Ajouter un soulignement au titre
        },
      },
      pageOrientation: 'landscape',
    };

    pdfMake.vfs = pdfFonts.pdfMake.vfs;
    const pdfDoc = pdfMake.createPdf(documentDefinition);

    pdfDoc.download('rapport_acces_employes.pdf');
  }

  getBase64ImageFromURL(url: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        if (ctx) {
          canvas.width = img.width;
          canvas.height = img.height;
          ctx.drawImage(img, 0, 0);
          const dataURL = canvas.toDataURL('image/png');
          resolve(dataURL);
        } else {
          reject(new Error("Impossible d'obtenir le contexte 2D du canevas."));
        }
      };
      img.onerror = (error) => {
        reject(error);
      };
      img.src = url;
    });
  }

  resetLocalStorage() {
    // Afficher une boîte de dialogue de confirmation
    Swal.fire({
      title: 'Êtes-vous sûr?',
      text: 'Cette action réinitialisera le stockage local, êtes-vous sûr de vouloir continuer?',
      icon: 'warning',
      showCancelButton: true,
      confirmButtonText: 'Oui, réinitialiser!',
      cancelButtonText: 'Annuler',
    }).then((result) => {
      if (result.isConfirmed) {
        for (let key in localStorage) {
          // Si l'utilisateur confirme, réinitialiser le stockage local
          if (key !== 'employees') {
            localStorage.removeItem(key);
          }
        }
        // Afficher une boîte de dialogue de succès
        Swal.fire({
          title: 'Réinitialisation réussie!',
          text: 'Le stockage local a été réinitialisé avec succès.',
          icon: 'success',
        });
      }
    });
  }
}
