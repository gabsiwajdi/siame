import { Component, OnInit } from '@angular/core';
import Swal from 'sweetalert2';
import { HostListener } from '@angular/core';
import { Router } from '@angular/router';

//** import pdf maker */
const pdfMake = require('pdfmake/build/pdfmake');
const pdfFonts = require('pdfmake/build/vfs_fonts');
pdfMake.vfs = pdfFonts.pdfMake.vfs;

@Component({
  selector: 'app-authentification',
  templateUrl: './authentification.component.html',
  styleUrls: ['./authentification.component.scss'],
})
export class AuthentificationComponent implements OnInit {
  enteredPassword: string = '';
  constructor(private router: Router) {}

  ngOnInit(): void {}

  @HostListener('window:keyup', ['$event'])
  keyEvent(event: KeyboardEvent) {
    // Vérifier si la touche Entrée a été enfoncée
    if (event.key === 'Enter') {
      // Vérifier si le mot de passe est 'infonetsiame'
      if (this.enteredPassword === 'infonetsiame') {
        // Rediriger vers le composant Configuration
        this.router.navigate(['/configuration']);
      } else {
        // Vérifier le mot de passe saisi
        this.checkEnteredPassword();
      }
    } else {
      // Ajouter la touche à la saisie du mot de passe
      this.enteredPassword += event.key;
    }
  }

  checkEnteredPassword() {
    const enteredCode = this.enteredPassword;

    // Récupérer les employés du stockage local
    const storedEmployees = localStorage.getItem('employees');
    const employees = storedEmployees ? JSON.parse(storedEmployees) : [];

    // Trouver l'employé correspondant au code entré
    const employeeMatch = employees.find(
      (employee: any) => employee.code === enteredCode
    );

    if (employeeMatch) {
      const today = new Date().toLocaleDateString(); // Date actuelle sous forme de chaîne (sans l'heure)

      // Vérifier si le code a déjà été utilisé aujourd'hui
      let usedCodes: string[] = [];
      const storedCodes = localStorage.getItem(today);
      if (storedCodes) {
        usedCodes = JSON.parse(storedCodes);
      }
      if (usedCodes.includes(enteredCode)) {
        // Affichage de la notification d'erreur si le code a déjà été utilisé aujourd'hui
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: "Tu as déjà utilisé ton ticket resto pour aujourd'hui",
          timer: 1500,
        });
      } else {
        // Vérifier l'autorisation de l'employé
        if (employeeMatch.autorisation === 'oui') {
          // Génération du ticket de restauration et impression
          const employeeName = employeeMatch.nom; // Utiliser le nom de l'employé correspondant
          const ticket = this.generateAndPrintTicket(employeeName);
          // Imprimer le ticket (vous devrez implémenter cette fonction)

          // Enregistrement du code utilisé dans le local storage pour aujourd'hui
          usedCodes.push(enteredCode);
          localStorage.setItem(today, JSON.stringify(usedCodes));
        } else {
          // Affichage de la notification d'erreur si l'employé n'est pas autorisé
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: "Vous n'êtes pas autorisé à utiliser le ticket resto",
            timer: 1500,
          });
        }
      }
    } else {
      // Affichage de la notification d'erreur si le code entré ne correspond à aucun code d'employé
      Swal.fire({
        icon: 'error',
        title: 'Oops...',
        text: "Le code entré est incorrect ou n'existe pas",
        timer: 1500,
      });
    }

    // Remise à zéro de la saisie du mot de passe
    this.enteredPassword = '';
  }

  generateAndPrintTicket(employeeName: string) {
    // Obtenez la date et l'heure actuelles
    const currentDate = new Date();
    const formattedDateTime = currentDate.toLocaleString(); // Date et heure au format complet

    // Récupérer le dernier numéro de ticket stocké dans le local storage
    let lastTicketNumber = localStorage.getItem('lastTicketNumber');
    let ticketNumber: number;

    // Vérifier s'il y a un dernier numéro de ticket stocké
    if (lastTicketNumber) {
      // Convertir le dernier numéro de ticket en nombre
      ticketNumber = parseInt(lastTicketNumber, 10) + 1;
    } else {
      // Si aucun numéro de ticket n'a été stocké, commencer par 1
      ticketNumber = 1;
    }

    // Mettre à jour le dernier numéro de ticket dans le local storage
    localStorage.setItem('lastTicketNumber', ticketNumber.toString());

    // Créer une nouvelle image
    const image = new Image();

    // Écouter l'événement load de l'image
    image.onload = () => {
      // Construire le contenu du ticket au format HTML
      const ticketContent = `
            <div style="text-align: center; margin-bottom: 0;">
                <img src="../../assets/images/SIAME.jpg" alt="Photo" style="width: 150px; height: auto;">
                <h2>N° : ${ticketNumber}</h2>
                <p>Date/Heure : ${formattedDateTime}</p>
                <p>Nom : ${employeeName}</p>
            </div>
        `;

      // Construire la fenêtre pour afficher le ticket au format HTML
      const newWindow = window.open('', '_blank');

      // Vérifier si la fenêtre a été ouverte avec succès
      if (newWindow) {
        // Injecter le contenu HTML dans la nouvelle fenêtre
        newWindow.document.write(ticketContent);

        // Appeler directement la fonction d'impression
        newWindow.focus(); // Mettre la fenêtre du ticket en focus
        newWindow.print(); // Appeler la fonction d'impression

        // Fermer le flux de document pour que le navigateur puisse terminer le rendu
        newWindow.document.close();

        // Définir un délai pour fermer la fenêtre après l'impression
        setTimeout(() => {
          newWindow.close();
        }, 1500); // Fermer la fenêtre après 1 seconde (ajustez selon vos besoins)
      } else {
        // Afficher un message d'erreur si la fenêtre n'a pas pu être ouverte
        console.error(
          "Impossible d'ouvrir une nouvelle fenêtre pour afficher le ticket."
        );
      }
    };

    // Définir la source de l'image
    image.src = '../../assets/images/SIAME.jpg';
  }
}
