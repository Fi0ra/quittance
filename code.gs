{\rtf1\ansi\ansicpg1252\cocoartf2821
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red20\green67\blue174;\red246\green247\blue249;\red46\green49\blue51;
\red24\green25\blue27;\red186\green6\blue115;\red162\green0\blue16;\red77\green80\blue85;\red18\green115\blue126;
\red97\green3\blue173;}
{\*\expandedcolortbl;;\cssrgb\c9412\c35294\c73725;\cssrgb\c97255\c97647\c98039;\cssrgb\c23529\c25098\c26275;
\cssrgb\c12549\c12941\c14118;\cssrgb\c78824\c15294\c52549;\cssrgb\c70196\c7843\c7059;\cssrgb\c37255\c38824\c40784;\cssrgb\c3529\c52157\c56863;
\cssrgb\c46275\c15294\c73333;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs26 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 genererQuittances\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 getActiveSpreadsheet\cf4 \strokec4 ();\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 configFeuille\cf4 \strokec4  = \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf7 \strokec7 "Configuration"\cf4 \strokec4 ); \cf8 \strokec8 // Assuming you named it "Configuration"\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 paiementFeuille\cf4 \strokec4  = \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "Paiements"\cf4 \strokec4 ));\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 locataireFeuille\cf4 \strokec4  = \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "Locataire"\cf4 \strokec4 ));\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 quittanceModeleFeuille\cf4 \strokec4  = \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "ModeleQuittance"\cf4 \strokec4 ));\cb1 \
\
\cb3   \cf2 \strokec2 if\cf4 \strokec4  (!\cf5 \strokec5 configFeuille\cf4 \strokec4  || !\cf5 \strokec5 paiementFeuille\cf4 \strokec4  || !\cf5 \strokec5 locataireFeuille\cf4 \strokec4  || !\cf5 \strokec5 quittanceModeleFeuille\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Error: One or more sheets not found. Check 'Configuration' sheet names."\cf4 \strokec4 );\cb1 \
\cb3     \cf2 \strokec2 return\cf4 \strokec4 ; \cf8 \strokec8 // Stop the script if sheets are missing\cf4 \cb1 \strokec4 \
\cb3   \}\cb1 \
\
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 paiementDonnees\cf4 \strokec4  = \cf5 \strokec5 paiementFeuille\cf4 \strokec4 .\cf5 \strokec5 getDataRange\cf4 \strokec4 ().\cf5 \strokec5 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 locataireDonnees\cf4 \strokec4  = \cf5 \strokec5 locataireFeuille\cf4 \strokec4 .\cf5 \strokec5 getDataRange\cf4 \strokec4 ().\cf5 \strokec5 getValues\cf4 \strokec4 ();\cb1 \
\
\cb3   \cf8 \strokec8 // Assuming these column numbers - ADJUST THESE BASED ON YOUR SHEET!\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_MOIS_ANNEE_COL\cf4 \strokec4  = \cf9 \strokec9 1\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_LOYER_COL\cf4 \strokec4  = \cf9 \strokec9 2\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_CHARGES_COL\cf4 \strokec4  = \cf9 \strokec9 3\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_LOCATAIRE_COL\cf4 \strokec4  = \cf9 \strokec9 4\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_QUITTANCE_SENT_COL\cf4 \strokec4  = \cf9 \strokec9 5\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 PAIEMENT_DATE_PAIEMENT_COL\cf4 \strokec4  = \cf9 \strokec9 6\cf4 \strokec4 ;\cb1 \
\
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 LOCATAIRE_NOM_COL\cf4 \strokec4  = \cf9 \strokec9 1\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 LOCATAIRE_ADRESSE_COL\cf4 \strokec4  = \cf9 \strokec9 2\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf6 \strokec6 LOCATAIRE_EMAIL_COL\cf4 \strokec4  = \cf9 \strokec9 3\cf4 \strokec4 ;\cb1 \
\
\cb3   \cf8 \strokec8 // Get configuration values\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 bailleurNom\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "BailleurNom"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 bailleurAdresse\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "BailleurAdresse"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 emailSignature\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "EmailSignature"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 quittancePrefix\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "QuittancePrefix"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 emailSubject\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "EmailSubject"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 emailBodyIntro\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "EmailBodyIntro"\cf4 \strokec4 ).\cf5 \strokec5 replace\cf4 \strokec4 (\cf10 \strokec10 /\\\\n/\cf2 \strokec2 g\cf4 \strokec4 , \cf7 \strokec7 "\\n"\cf4 \strokec4 ); \cf8 \strokec8 // Replace literal \\n with newline\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 emailBodyOutro\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "EmailBodyOutro"\cf4 \strokec4 ).\cf5 \strokec5 replace\cf4 \strokec4 (\cf10 \strokec10 /\\\\n/\cf2 \strokec2 g\cf4 \strokec4 , \cf7 \strokec7 "\\n"\cf4 \strokec4 ); \cf8 \strokec8 // Replace literal \\n with newline\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 dateFormatMois\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "DateFormatMois"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 dateFormatOptions\cf4 \strokec4  = \cf6 \strokec6 JSON\cf4 \strokec4 .\cf5 \strokec5 parse\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "DateFormatOptions"\cf4 \strokec4 )); \cf8 \strokec8 // Parse string to object\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 dateFormatAnnee\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "DateFormatAnnee"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 timeZone\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "TimeZone"\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 driveFolderId\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "DriveFolderID"\cf4 \strokec4 ); \cf8 \strokec8 // Get the Google Drive folder ID\cf4 \cb1 \strokec4 \
\
\
\
\cb3   \cf8 \strokec8 // Get the declaration template from configuration\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 quittanceDeclarationTemplate\cf4 \strokec4  = \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "QuittanceDeclaration"\cf4 \strokec4 );\cb1 \
\
\cb3   \cf8 \strokec8 // Pre-process locataire data for efficient lookup\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 locataires\cf4 \strokec4  = \{\};\cb1 \
\cb3   \cf2 \strokec2 for\cf4 \strokec4  (\cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 j\cf4 \strokec4  = \cf9 \strokec9 1\cf4 \strokec4 ; \cf5 \strokec5 j\cf4 \strokec4  < \cf5 \strokec5 locataireDonnees\cf4 \strokec4 .\cf5 \strokec5 length\cf4 \strokec4 ; \cf5 \strokec5 j\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf5 \strokec5 locataires\cf4 \strokec4 [\cf5 \strokec5 locataireDonnees\cf4 \strokec4 [\cf5 \strokec5 j\cf4 \strokec4 ][\cf6 \strokec6 LOCATAIRE_NOM_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ]] = \cf5 \strokec5 locataireDonnees\cf4 \strokec4 [\cf5 \strokec5 j\cf4 \strokec4 ];\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf2 \strokec2 for\cf4 \strokec4  (\cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 i\cf4 \strokec4  = \cf9 \strokec9 1\cf4 \strokec4 ; \cf5 \strokec5 i\cf4 \strokec4  < \cf5 \strokec5 paiementDonnees\cf4 \strokec4 .\cf5 \strokec5 length\cf4 \strokec4 ; \cf5 \strokec5 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf2 \strokec2 try\cf4 \strokec4  \{\cb1 \
\cb3       \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_QUITTANCE_SENT_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ] !== \cf2 \strokec2 true\cf4 \strokec4  && \cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_DATE_PAIEMENT_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ] !== \cf7 \strokec7 ""\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 moisAnnee\cf4 \strokec4  = \cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_MOIS_ANNEE_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 loyer\cf4 \strokec4  = \cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_LOYER_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 charges\cf4 \strokec4  = \cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_CHARGES_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 locataireNom\cf4 \strokec4  = \cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_LOCATAIRE_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 dateDePaiement\cf4 \strokec4  = \cf2 \strokec2 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf5 \strokec5 paiementDonnees\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf6 \strokec6 PAIEMENT_DATE_PAIEMENT_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ]);\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 dateFormatter\cf4 \strokec4  = \cf2 \strokec2 new\cf4 \strokec4  \cf6 \strokec6 Intl\cf4 \strokec4 .\cf6 \strokec6 DateTimeFormat\cf4 \strokec4 (\cf5 \strokec5 dateFormatMois\cf4 \strokec4 , \cf5 \strokec5 dateFormatOptions\cf4 \strokec4 );\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 mois\cf4 \strokec4  = \cf5 \strokec5 dateFormatter\cf4 \strokec4 .\cf5 \strokec5 format\cf4 \strokec4 (\cf5 \strokec5 dateDePaiement\cf4 \strokec4 );\cb1 \
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 annee\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf5 \strokec5 formatDate\cf4 \strokec4 (\cf5 \strokec5 dateDePaiement\cf4 \strokec4 , \cf5 \strokec5 timeZone\cf4 \strokec4 , \cf5 \strokec5 dateFormatAnnee\cf4 \strokec4 );\cb1 \
\
\cb3         \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 locataireInfos\cf4 \strokec4  = \cf5 \strokec5 locataires\cf4 \strokec4 [\cf5 \strokec5 locataireNom\cf4 \strokec4 ];\cb1 \
\
\cb3         \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 locataireInfos\cf4 \strokec4 ) \{\cb1 \
\cb3           \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 quittanceFeuille\cf4 \strokec4  = \cf5 \strokec5 quittanceModeleFeuille\cf4 \strokec4 .\cf5 \strokec5 copyTo\cf4 \strokec4 (\cf5 \strokec5 ss\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 setName\cf4 \strokec4 (\cf5 \strokec5 quittancePrefix\cf4 \strokec4  + \cf7 \strokec7 " "\cf4 \strokec4  + \cf5 \strokec5 moisAnnee\cf4 \strokec4 );\cb1 \
\
\cb3           \cf8 \strokec8 // *** ADJUST THESE RANGES BASED ON YOUR "Modele Quittance" TAB ***\cf4 \cb1 \strokec4 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B2"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 moisAnnee\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B3"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 locataireInfos\cf4 \strokec4 [\cf6 \strokec6 LOCATAIRE_ADRESSE_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ]);\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "A5"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 bailleurNom\cf4 \strokec4 ); \cf8 \strokec8 // Bailleur Name from config\cf4 \cb1 \strokec4 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B17"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 bailleurAdresse\cf4 \strokec4 ); \cf8 \strokec8 // Bailleur Address from config\cf4 \cb1 \strokec4 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "C5"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 locataireInfos\cf4 \strokec4 [\cf6 \strokec6 LOCATAIRE_NOM_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ]);\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B7"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 loyer\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B8"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 charges\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B9"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 loyer\cf4 \strokec4  + \cf5 \strokec5 charges\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B11"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 dateDePaiement\cf4 \strokec4 );\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "B18"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf2 \strokec2 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ());\cb1 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "C18"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 bailleurNom\cf4 \strokec4 );\cb1 \
\cb3           \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 quittanceDeclaration\cf4 \strokec4  = \cf5 \strokec5 quittanceDeclarationTemplate\cf4 \cb1 \strokec4 \
\cb3             .\cf5 \strokec5 replace\cf4 \strokec4 (\cf10 \strokec10 /\\\{BailleurNom\\\}/\cf2 \strokec2 g\cf4 \strokec4 , \cf5 \strokec5 bailleurNom\cf4 \strokec4 ) \cf8 \strokec8 // Replace placeholder\cf4 \cb1 \strokec4 \
\cb3             .\cf5 \strokec5 replace\cf4 \strokec4 (\cf10 \strokec10 /\\\\n/\cf2 \strokec2 g\cf4 \strokec4 , \cf7 \strokec7 "\\n"\cf4 \strokec4 ); \cf8 \strokec8 // Replace literal \\n with newline\cf4 \cb1 \strokec4 \
\cb3           \cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf7 \strokec7 "A13"\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf5 \strokec5 quittanceDeclaration\cf4 \strokec4 ); \cf8 \strokec8 // Adjust "A15" to your cell\cf4 \cb1 \strokec4 \
\cb3         \cb1 \
\
\cb3           \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 flush\cf4 \strokec4 (); \cf8 \strokec8 // Force write operations to complete\cf4 \cb1 \strokec4 \
\
\cb3           \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 pdf\cf4 \strokec4  = \cf5 \strokec5 genererPDFWithRetry\cf4 \strokec4 (\cf5 \strokec5 quittanceFeuille\cf4 \strokec4 .\cf5 \strokec5 getSheetId\cf4 \strokec4 (), \cf5 \strokec5 configFeuille\cf4 \strokec4 ); \cf8 \strokec8 // Pass configFeuille\cf4 \cb1 \strokec4 \
\
\cb3           \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 pdf\cf4 \strokec4 ) \{\cb1 \
\cb3             \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 pdfFilename\cf4 \strokec4  = \cf5 \strokec5 quittancePrefix\cf4 \strokec4  + \cf7 \strokec7 " "\cf4 \strokec4  + \cf5 \strokec5 moisAnnee\cf4 \strokec4  + \cf7 \strokec7 ".pdf"\cf4 \strokec4 ;\cb1 \
\cb3             \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 pieceJointe\cf4 \strokec4  = \{\cb1 \
\cb3               \cf5 \strokec5 fileName\cf4 \strokec4 : \cf5 \strokec5 pdfFilename\cf4 \strokec4 ,\cb1 \
\cb3               \cf5 \strokec5 content\cf4 \strokec4 : \cf5 \strokec5 pdf\cf4 \strokec4 ,\cb1 \
\cb3               \cf5 \strokec5 mimeType\cf4 \strokec4 : \cf7 \strokec7 "application/pdf"\cf4 \strokec4 ,\cb1 \
\cb3             \};\cb1 \
\
\cb3             \cf8 \strokec8 // Save to Google Drive\cf4 \cb1 \strokec4 \
\cb3             \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 driveFolderId\cf4 \strokec4 ) \{\cb1 \
\cb3               \cf2 \strokec2 try\cf4 \strokec4  \{\cb1 \
\cb3                 \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 folder\cf4 \strokec4  = \cf6 \strokec6 DriveApp\cf4 \strokec4 .\cf5 \strokec5 getFolderById\cf4 \strokec4 (\cf5 \strokec5 driveFolderId\cf4 \strokec4 );\cb1 \
\cb3                 \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 pdfBlob\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf5 \strokec5 newBlob\cf4 \strokec4 (\cf5 \strokec5 pdf\cf4 \strokec4 , \cf6 \strokec6 MimeType\cf4 \strokec4 .\cf6 \strokec6 PDF\cf4 \strokec4 , \cf5 \strokec5 pdfFilename\cf4 \strokec4 );\cb1 \
\cb3                 \cf5 \strokec5 folder\cf4 \strokec4 .\cf5 \strokec5 createFile\cf4 \strokec4 (\cf5 \strokec5 pdfBlob\cf4 \strokec4 );\cb1 \
\cb3                 \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "PDF saved to Drive: "\cf4 \strokec4  + \cf5 \strokec5 pdfFilename\cf4 \strokec4 );\cb1 \
\cb3               \} \cf2 \strokec2 catch\cf4 \strokec4  (\cf5 \strokec5 e\cf4 \strokec4 ) \{\cb1 \
\cb3                 \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Error saving PDF to Drive: "\cf4 \strokec4  + \cf5 \strokec5 e\cf4 \strokec4 .\cf5 \strokec5 toString\cf4 \strokec4 ());\cb1 \
\cb3               \}\cb1 \
\cb3             \} \cf2 \strokec2 else\cf4 \strokec4  \{\cb1 \
\cb3               \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "DriveFolderID not configured. PDF not saved to Drive."\cf4 \strokec4 );\cb1 \
\cb3             \}\cb1 \
\
\cb3             \cf6 \strokec6 MailApp\cf4 \strokec4 .\cf5 \strokec5 sendEmail\cf4 \strokec4 (\{\cb1 \
\cb3               \cf5 \strokec5 to\cf4 \strokec4 : \cf5 \strokec5 locataireInfos\cf4 \strokec4 [\cf6 \strokec6 LOCATAIRE_EMAIL_COL\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ],\cb1 \
\cb3               \cf5 \strokec5 subject\cf4 \strokec4 : \cf5 \strokec5 emailSubject\cf4 \strokec4  + \cf7 \strokec7 " "\cf4 \strokec4  + \cf5 \strokec5 moisAnnee\cf4 \strokec4 ,\cb1 \
\cb3               \cf5 \strokec5 body\cf4 \strokec4 : \cf5 \strokec5 emailBodyIntro\cf4 \strokec4  + \cf7 \strokec7 " "\cf4 \strokec4  + \cf5 \strokec5 moisAnnee\cf4 \strokec4  + \cf5 \strokec5 emailBodyOutro\cf4 \strokec4  + \cf5 \strokec5 emailSignature\cf4 \strokec4 ,\cb1 \
\cb3               \cf5 \strokec5 attachments\cf4 \strokec4 : [\cf5 \strokec5 pieceJointe\cf4 \strokec4 ],\cb1 \
\cb3             \});\cb1 \
\
\cb3             \cf5 \strokec5 paiementFeuille\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf5 \strokec5 i\cf4 \strokec4  + \cf9 \strokec9 1\cf4 \strokec4 , \cf6 \strokec6 PAIEMENT_QUITTANCE_SENT_COL\cf4 \strokec4 ).\cf5 \strokec5 setValue\cf4 \strokec4 (\cf2 \strokec2 true\cf4 \strokec4 );\cb1 \
\cb3             \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 flush\cf4 \strokec4 (); \cf8 \strokec8 // Force write operations to complete\cf4 \cb1 \strokec4 \
\cb3           \} \cf2 \strokec2 else\cf4 \strokec4  \{\cb1 \
\cb3             \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Error generating PDF for: "\cf4 \strokec4  + \cf5 \strokec5 moisAnnee\cf4 \strokec4 );\cb1 \
\cb3           \}\cb1 \
\
\cb3           \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 deleteSheet\cf4 \strokec4 (\cf5 \strokec5 quittanceFeuille\cf4 \strokec4 ); \cf8 \strokec8 // Clean up the temporary sheet\cf4 \cb1 \strokec4 \
\cb3         \} \cf2 \strokec2 else\cf4 \strokec4  \{\cb1 \
\cb3           \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Locataire not found: "\cf4 \strokec4  + \cf5 \strokec5 locataireNom\cf4 \strokec4 );\cb1 \
\cb3         \}\cb1 \
\cb3       \}\cb1 \
\cb3     \} \cf2 \strokec2 catch\cf4 \strokec4  (\cf5 \strokec5 e\cf4 \strokec4 ) \{\cb1 \
\cb3       \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Error processing payment for "\cf4 \strokec4  + \cf5 \strokec5 e\cf4 \strokec4 .\cf5 \strokec5 toString\cf4 \strokec4 ());\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configSheet\cf4 \strokec4 , \cf5 \strokec5 settingName\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 data\cf4 \strokec4  = \cf5 \strokec5 configSheet\cf4 \strokec4 .\cf5 \strokec5 getDataRange\cf4 \strokec4 ().\cf5 \strokec5 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf2 \strokec2 for\cf4 \strokec4  (\cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 i\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 ; \cf5 \strokec5 i\cf4 \strokec4  < \cf5 \strokec5 data\cf4 \strokec4 .\cf5 \strokec5 length\cf4 \strokec4 ; \cf5 \strokec5 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 data\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ] === \cf5 \strokec5 settingName\cf4 \strokec4 ) \{\cb1 \
\cb3       \cf2 \strokec2 return\cf4 \strokec4  \cf5 \strokec5 data\cf4 \strokec4 [\cf5 \strokec5 i\cf4 \strokec4 ][\cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\cb3   \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Warning: Setting '"\cf4 \strokec4  + \cf5 \strokec5 settingName\cf4 \strokec4  + \cf7 \strokec7 "' not found in Configuration sheet."\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 return\cf4 \strokec4  \cf2 \strokec2 null\cf4 \strokec4 ; \cf8 \strokec8 // Or throw an error: throw new Error("Setting '" + settingName + "' not found...");\cf4 \cb1 \strokec4 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 genererPDFWithRetry\cf4 \strokec4 (\cf5 \strokec5 quittanceFeuilleId\cf4 \strokec4 , \cf5 \strokec5 configFeuille\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 maxRetries\cf4 \strokec4  = \cf5 \strokec5 parseInt\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "PDFRetryAttempts"\cf4 \strokec4 )) || \cf9 \strokec9 3\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 retryDelay\cf4 \strokec4  = \cf5 \strokec5 parseInt\cf4 \strokec4 (\cf5 \strokec5 getConfigValue\cf4 \strokec4 (\cf5 \strokec5 configFeuille\cf4 \strokec4 , \cf7 \strokec7 "PDFRetryDelay"\cf4 \strokec4 )) || \cf9 \strokec9 2000\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 attempt\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 ;\cb1 \
\cb3   \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 pdf\cf4 \strokec4  = \cf2 \strokec2 null\cf4 \strokec4 ;\cb1 \
\
\cb3   \cf2 \strokec2 while\cf4 \strokec4  (\cf5 \strokec5 attempt\cf4 \strokec4  < \cf5 \strokec5 maxRetries\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf2 \strokec2 try\cf4 \strokec4  \{\cb1 \
\cb3       \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 getActiveSpreadsheet\cf4 \strokec4 ();\cb1 \
\cb3       \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 sheet\cf4 \strokec4  = \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getSheets\cf4 \strokec4 ().\cf5 \strokec5 find\cf4 \strokec4 (\cf2 \strokec2 function\cf4 \strokec4 (\cf5 \strokec5 sheet\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf2 \strokec2 return\cf4 \strokec4  \cf5 \strokec5 sheet\cf4 \strokec4 .\cf5 \strokec5 getSheetId\cf4 \strokec4 () === \cf5 \strokec5 quittanceFeuilleId\cf4 \strokec4 ;\cb1 \
\cb3       \});\cb1 \
\
\cb3       \cf2 \strokec2 if\cf4 \strokec4  (!\cf5 \strokec5 sheet\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Sheet not found with ID: "\cf4 \strokec4  + \cf5 \strokec5 quittanceFeuilleId\cf4 \strokec4 );\cb1 \
\cb3         \cf2 \strokec2 return\cf4 \strokec4  \cf2 \strokec2 null\cf4 \strokec4 ;\cb1 \
\cb3       \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf8 \cb3 \strokec8 // Generate a PDF from the sheet with specific parameters for content size\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3       \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 url\cf4 \strokec4  = \cf7 \strokec7 "https://docs.google.com/spreadsheets/d/"\cf4 \strokec4  + \cf5 \strokec5 ss\cf4 \strokec4 .\cf5 \strokec5 getId\cf4 \strokec4 () + \cf7 \strokec7 "/export?"\cf4 \strokec4  +\cb1 \
\cb3                 \cf7 \strokec7 "format=pdf"\cf4 \strokec4  +                     \cf8 \strokec8 // Use format=pdf\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&gid="\cf4 \strokec4  + \cf5 \strokec5 sheet\cf4 \strokec4 .\cf5 \strokec5 getSheetId\cf4 \strokec4 () +\cb1 \
\cb3                 \cf7 \strokec7 "&size=A5"\cf4 \strokec4  +                       \cf8 \strokec8 // Use a standard size like A4 or letter\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&portrait=true"\cf4 \strokec4  +                 \cf8 \strokec8 // Or false if landscape\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&scale=4"\cf4 \strokec4  +                       \cf8 \strokec8 // IMPORTANT: Use scale=1 for 100% size\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&top_margin=0.5"\cf4 \strokec4  +                  \cf8 \strokec8 // IMPORTANT: Remove margins\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&bottom_margin=0.5"\cf4 \strokec4  +\cb1 \
\cb3                 \cf7 \strokec7 "&left_margin=0.5"\cf4 \strokec4  +\cb1 \
\cb3                 \cf7 \strokec7 "&right_margin=0.5"\cf4 \strokec4  +\cb1 \
\cb3                 \cf7 \strokec7 "&gridlines=false"\cf4 \strokec4  +               \cf8 \strokec8 // Optional: Hide gridlines\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&printnotes=false"\cf4 \strokec4  +              \cf8 \strokec8 // Optional: Hide notes\cf4 \cb1 \strokec4 \
\cb3                 \cf7 \strokec7 "&fzr=false"\cf4 \strokec4 ;                      \cf8 \strokec8 // Optional: Prevent row zoom fit\cf4 \cb1 \strokec4 \
\
\cb3       \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 params\cf4 \strokec4  = \{\cb1 \
\cb3         \cf5 \strokec5 method\cf4 \strokec4 : \cf7 \strokec7 "get"\cf4 \strokec4 ,\cb1 \
\cb3         \cf5 \strokec5 headers\cf4 \strokec4 : \{ \cf7 \strokec7 "Authorization"\cf4 \strokec4 : \cf7 \strokec7 "Bearer "\cf4 \strokec4  + \cf6 \strokec6 ScriptApp\cf4 \strokec4 .\cf5 \strokec5 getOAuthToken\cf4 \strokec4 () \}\cb1 \
\cb3       \};\cb1 \
\
\cb3       \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 blob\cf4 \strokec4  = \cf6 \strokec6 UrlFetchApp\cf4 \strokec4 .\cf5 \strokec5 fetch\cf4 \strokec4 (\cf5 \strokec5 url\cf4 \strokec4 , \cf5 \strokec5 params\cf4 \strokec4 ).\cf5 \strokec5 getBlob\cf4 \strokec4 ().\cf5 \strokec5 setContentType\cf4 \strokec4 (\cf7 \strokec7 "application/pdf"\cf4 \strokec4 );\cb1 \
\
\cb3       \cf5 \strokec5 pdf\cf4 \strokec4  = \cf5 \strokec5 blob\cf4 \strokec4 .\cf5 \strokec5 getBytes\cf4 \strokec4 ();\cb1 \
\cb3       \cf2 \strokec2 return\cf4 \strokec4  \cf5 \strokec5 pdf\cf4 \strokec4 ; \cf8 \strokec8 // Success, return the pdf\cf4 \cb1 \strokec4 \
\cb3     \} \cf2 \strokec2 catch\cf4 \strokec4  (\cf5 \strokec5 e\cf4 \strokec4 ) \{\cb1 \
\cb3       \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "Attempt "\cf4 \strokec4  + (\cf5 \strokec5 attempt\cf4 \strokec4  + \cf9 \strokec9 1\cf4 \strokec4 ) + \cf7 \strokec7 " failed: "\cf4 \strokec4  + \cf5 \strokec5 e\cf4 \strokec4 .\cf5 \strokec5 toString\cf4 \strokec4 ());\cb1 \
\cb3       \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf5 \strokec5 sleep\cf4 \strokec4 (\cf5 \strokec5 retryDelay\cf4 \strokec4 ); \cf8 \strokec8 // Wait before retrying\cf4 \cb1 \strokec4 \
\cb3       \cf5 \strokec5 attempt\cf4 \strokec4 ++;\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf7 \strokec7 "PDF generation failed after "\cf4 \strokec4  + \cf5 \strokec5 maxRetries\cf4 \strokec4  + \cf7 \strokec7 " attempts."\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 return\cf4 \strokec4  \cf2 \strokec2 null\cf4 \strokec4 ; \cf8 \strokec8 // PDF generation failed\cf4 \cb1 \strokec4 \
\cb3 \}\cb1 \
}