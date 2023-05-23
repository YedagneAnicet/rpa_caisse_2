*** Settings ***
Documentation     Automatisation du processus d'insertion de donnée dans une base d'un fichier excel
Suite Setup       Connect To Database    psycopg2    ${DBName}    ${DBUser}    ${DBPass}    ${DBHost}    ${DBPort}
Suite Teardown    Disconnect From Database
Library           DatabaseLibrary
Library           ExcelLibrary
Library           BuiltIn
Library           SeleniumLibrary
Library           DateTime
Library           RPA.Email.ImapSmtp   smtp_server=smtp.gmail.com  smtp_port=587
Library           RPA.PDF
Library           String
Library           Collections
Task Setup  Authorize  account=${gmail}  password=${mdp}

*** Variables ***
${gmail}          testpython.eburtis@gmail.com
${mdp}            vfrjkmwvxfcrbdfr

${DBHost}         localhost
${DBName}         rpa_caisse_db
${DBPass}         postgres
${DBPort}         5432
${DBUser}         postgres

${chemin}         ${CURDIR}${/}\\ressources\\donneecaisse.xlsx
${feuille}        rapport



*** Keywords ***
Lire le fichier Excel
    [Documentation]        Lire le fichier excel et recuperer les données neccessaires
    [Arguments]                     ${chemin}       ${feuille}      ${ligne}    ${colonne}
    Open Excel Document             ${chemin}           1
    Get Sheet                       ${feuille}
    ${value}    Read Excel Cell     ${ligne}        ${colonne}
    [Return]                        ${value}
    Close All Excel Documents

Insertion dans la base de donnee
    [Documentation]      Insertion des données dans la base de donnée
    [Arguments]    ${nom_responsable}    ${email_responsable}    ${date}    ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${query}    Catenate       INSERT INTO  rapport_journalier (nom_responsable, email_responsable, date, montant_carte_bancaire, montant_espece, montant_ticket_restaurant, montant_prelevement, montant_monnaie ) VALUES ('${nom_responsable}','${email_responsable}','${date}','${montant_carte_bancaire}','${montant_espece}','${montant_ticket_restaurant}','${montant_prelevement}','${montant_apport_monnaie}')
    Execute Sql String    ${query}

Recuperer les donnee de la base de donnee
    [Documentation]    Recuperation des données enregistrées dans la base de donneecaisse
    ${query}           Catenate              SELECT * FROM rapport_journalier
    @{donnee}          Query    ${query}
    [Return]           @{donnee}


Vérification des montants
    [Documentation]     Vérification du solde selon la regles "carte bancaire + espèces + ticket restaurant = prélèvement - apport monnaie".
    [Arguments]         ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${montant_total}    Evaluate                    ${montant_carte_bancaire}+${montant_espece}+${montant_ticket_restaurant}
    ${solde}            Evaluate                    ${montant_prelevement} - ${montant_apport_monnaie}
    ${statut_solde}     Run Keyword If              '${montant_total}'=='${solde}'    Set Variable    ${True}     ELSE    Set Variable    ${False}
    [Return]            ${statut_solde}

Formattage Date
    [Documentation]     Fait le formatage de date selon le model 25/05/24
    [Arguments]         ${date}
    ${date}             Get Current Date    result_format=%d/%m/%Y
    [Return]            ${date}


Envoie de mail en cas d'erreur
    [Arguments]   ${email_responsable}       ${date}
    Send Message  sender=${gmail}
    ...           recipients=${email_responsable}
    ...           subject=RPA CAISSE
    ...           body=Bonjour, J'ai trouvé une erreur dans le rapport journalier du ${date} sur les montants, \n Merci de verifier les differents montants. \n Cordialement
    ...           attachments=${chemin}

*** Test Cases ***
Enregistrement du rapport dans la base de donnee
    [Documentation]        Recuperation des données du fichier excel et enregistrement dans la base de donnee ensuite effectue une verification en cas d'erreur envoie un mail
    ${nom_responsable}              Lire le fichier Excel       ${chemin}       ${feuille}       3       3
    ${email_responsable}            Lire le fichier Excel       ${chemin}       ${feuille}       4       3
    ${date}                         Lire le fichier Excel       ${chemin}       ${feuille}       5       3
    ${montant_carte_bancaire}       Lire le fichier Excel       ${chemin}       ${feuille}       11      4
    ${montant_espece}               Lire le fichier Excel       ${chemin}       ${feuille}       12      4
    ${montant_ticket_restaurant}    Lire le fichier Excel       ${chemin}       ${feuille}       13      4
    ${montant_prelevement}          Lire le fichier Excel       ${chemin}       ${feuille}       15      4
    ${montant_apport_monnaie}       Lire le fichier Excel       ${chemin}       ${feuille}       16      4

    Insertion dans la base de donnee        ${nom_responsable}          ${email_responsable}    ${date}    ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${status_solde}        Vérification des montants               ${montant_carte_bancaire}   ${montant_espece}       ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    
    ${date_format}      Formattage Date    ${date}

    IF    ${status_solde} == False
         Envoie de mail en cas d'erreur    ${email_responsable}    ${date_format}
    END
    
Rapport d'activité
    @{ListeRapports}    Recuperer les donnee de la base de donnee

    @{ResponsablesEnErreur}     Create List   

    FOR    ${rapport}    IN    @{ListeRapports}
        ${nom_responsable}               Set Variable        ${rapport}[1]
        ${email_responsable}             Set Variable        ${rapport}[2]
        ${date}                          Set Variable        ${rapport}[3]
        ${montant_carte_bancaire}        Set Variable        ${rapport}[4]
        ${montant_espece}                Set Variable        ${rapport}[5]
        ${montant_ticket_restaurant}     Set Variable        ${rapport}[6]
        ${montant_prelevement}           Set Variable        ${rapport}[7]
        ${montant_apport_monnaie}        Set Variable        ${rapport}[8]

        ${status_solde}        Vérification des montants               ${montant_carte_bancaire}   ${montant_espece}       ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}    
        
        ${date_format}        Formattage Date    ${date}
        
        IF    ${status_solde} == False  
            Append To List     ${ResponsablesEnErreur}    ${nom_responsable} 
        END
    
    END

    ${contenu_html}    Create List    

    FOR    ${responsable}    IN    @{ResponsablesEnErreur}
            ${balise_html}    Set Variable    <!DOCTYPE html><html><body> <h2>${date_format}</h2><ol><li>${responsable}</li></ol></body></html>
            Append To List    ${contenu_html}    ${balise_html}    
    END
    
    
    ${rapport_complet}    Set Variable    ${contenu_html}    
    
    Html To Pdf    ${rapport_complet}    test.pdf