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
Library           StringFormat
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
    [Arguments]    ${nom_responsable}    ${email_responsable}    ${date}    ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   
    ...    ${montant_prelevement}   ${montant_apport_monnaie}
    ${query}    Catenate       INSERT INTO  rapport_journalier (nom_responsable, email_responsable, date, montant_carte_bancaire, montant_espece, montant_ticket_restaurant,
    ...    montant_prelevement, montant_monnaie ) VALUES ('${nom_responsable}','${email_responsable}','${date}','${montant_carte_bancaire}','${montant_espece}',
    ...    '${montant_ticket_restaurant}','${montant_prelevement}','${montant_apport_monnaie}')
    Execute Sql String    ${query}

Recuperer les donnee de la base de donnee
    [Documentation]    Recuperation des données enregistrées dans la base de donneecaisse
    ${query}           Catenate              SELECT date, JSONB_AGG(JSONB_BUILD_OBJECT(
    ...    'nom_responsable', nom_responsable,
    ...    'email_responsable', email_responsable,
    ...    'montant_carte_bancaire', montant_carte_bancaire,
    ...    'montant_espece', montant_espece,
    ...    'montant_ticket_restaurant', montant_ticket_restaurant,
    ...    'montant_prelevement', montant_prelevement,
    ...    'montant_monnaie', montant_monnaie
    ...    )) AS rapports FROM rapport_journalier GROUP BY date;
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

    @{ResponsablesEnErreur}    Create List

    ${rapport_complet}    Set Variable    ${EMPTY}

    FOR    ${rapport}    IN    @{ListeRapports}

        ${date}    Set Variable    ${rapport[0]}
        ${rapport_details}    Set Variable    ${rapport[1]}

        FOR    ${responsable_details}    IN    @{rapport_details}
            ${nom_responsable}                Set Variable        ${responsable_details['nom_responsable']}
            ${email_responsable}              Set Variable        ${responsable_details['email_responsable']}
            ${montant_carte_bancaire}         Set Variable        ${responsable_details['montant_carte_bancaire']}
            ${montant_espece}                 Set Variable        ${responsable_details['montant_espece']}
            ${montant_ticket_restaurant}      Set Variable        ${responsable_details['montant_ticket_restaurant']}
            ${montant_prelevement}            Set Variable        ${responsable_details['montant_prelevement']}
            ${montant_monnaie}                Set Variable        ${responsable_details['montant_monnaie']}
            
            ${status_solde}    Vérification des montants    ${montant_carte_bancaire}    ${montant_espece}    ${montant_ticket_restaurant}    ${montant_prelevement}    ${montant_monnaie}
            
            IF    ${status_solde} == False
                Append To List    ${ResponsablesEnErreur}    ${nom_responsable}
            END
          
        END

        ${contenu_html}    Create List

        FOR    ${responsable}    IN    @{ResponsablesEnErreur}
            ${balise_html}    Catenate    <li>${responsable}</li>
            Append To List    ${contenu_html}    ${balise_html}
        END

        @{ResponsablesEnErreur}       Create List

        ${contenu_html}    Evaluate       ''.join(${contenu_html})    

        ${rapport_html}    Catenate    <h2>${date}</h2><ol>${contenu_html}</ol>

        ${rapport_complet}    Set Variable        ${rapport_complet}${rapport_html}
    END

    Html To Pdf    ${rapport_complet}    rapport.pdf

    Send Message  sender=${gmail}
    ...           recipients=romeo.beyara@eburtis.ci
    ...           subject=Rapport de caisse 
    ...           body=Bonjour, Ci-joint la liste des responsables dont les rapports journaliers sont en erreur , \n Merci de les contacter pour plus de details. \n Cordialement
    ...           attachments=${CURDIR}${/}rapport.pdf
