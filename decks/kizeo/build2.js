// Deck « Kizeo → Teepee » : identité propre (sombre/clair alterné, serif + sans, cercle mobile, barre de progression), PPTX natif + Morph.
const fs=require('fs'),path=require('path');const pptxgen=require('pptxgenjs');const JSZip=require('jszip');
const AS=process.argv[2],OUT=process.argv[3];const inch=v=>v/144,pt=v=>v/2;
const FH='Roboto',FB='Roboto';
const D={bg:'0E1B24',tx:'F4F1EA',mu:'9DB0BF',pn:'15283A',ln:'2B4457',ac:'3CC3B5',wm:'F26B3A'};
const L={bg:'F6F4EF',tx:'12202B',mu:'5A6A77',pn:'FFFFFF',ln:'D9D4C8',ac:'0F766E',wm:'D9480F'};
const SZ={'a1_cartographie.png':[1344,772],'a5_listes.png':[1224,616],'a5_formulaire.png':[1344,536],'a5_workflow.png':[630,292],'a2_teepee_mobile.png':[1344,740],'a2_guide.png':[1344,714],'gantt.png':[1330,590],'poster.png':[1920,1080]};
const S=[]; // slide builders
const def=(label,dark,orb,notes,fn)=>S.push({label,dark,orb,notes,fn});
const fit=(f,w,h)=>{const [a,b]=SZ[f],r=Math.min(w/a,h/b);return [a*r,b*r]};

// ---------- contenu ----------
def('Soutenance',true,[1050,380,1100],"Bonjour. Je présente le projet de transition numérique que j’ai mené chez VINCI Facilities Tertiaire Aquitaine : la migration des formulaires opérationnels de Kizeo vers Teepee, de septembre 2025 à la phase de stabilisation.",(c)=>{
 c.T(128,300,1300,40,'Soutenance · Projet de transition numérique',{size:28,color:c.k.ac,bold:1,sp:3,up:1});
 c.T(128,360,1500,200,'Kizeo  →  Teepee',{f:FH,size:150,bold:1});
 c.T(128,590,1150,140,'Migrer les formulaires opérationnels de VINCI Facilities Tertiaire Aquitaine',{size:44,color:c.k.mu,lh:1.2});
 c.T(128,800,1100,50,'Yann Yankam Tchapda',{size:40,bold:1});
 c.T(128,860,1300,80,'Ingénieur Systèmes Numériques et Génie Industriel · Icam · 2025–2026',{size:28,color:c.k.mu});
});
def('Le projet en bref',false,[1500,-250,700],"Avant d’entrer dans le détail, voici le projet en dix secondes : on part des formulaires Kizeo, on les rationalise, on les reconstruit dans Teepee et on forme les utilisateurs. Ma présentation suit quatre temps : le contexte, le pilotage, les résultats puis le bilan.",(c)=>{
 c.T(128,200,1500,100,'Le projet en bref',{f:FH,size:64,bold:1});
 [['01','Contexte','Une décision du groupe VINCI'],['02','Pilotage','Démarche, bugs, conduite du changement'],['03','Résultats','Livrables, chiffres et écarts'],['04','Bilan','Compétences, retour critique, perspectives']].forEach((r,i)=>{const y=330+i*150;
  c.T(128,y,130,100,r[0],{f:FH,size:64,color:c.k.ac,bold:1});c.T(270,y+4,520,56,r[1],{size:44,bold:1});c.T(270,y+62,560,40,r[2],{size:28,color:c.k.mu});c.R(128,y+122,700,2,{fill:c.k.ln});});
 c.R(868,300,924,520,{fill:c.k.pn,line:c.k.ln,lw:2,rad:0.12});
 c.V(880,312,900,506);
 c.T(868,838,924,40,'Vidéo : la migration de Kizeo vers Teepee en 10 secondes (cliquer pour lancer)',{size:24,color:c.k.mu,italic:1});
});
def('Contexte',true,[-250,500,600],"Le projet répond à une directive du groupe VINCI et à des exigences de sécurité : le contrat cadre Kizeo n’a pas été renouvelé, alors que Teepee est une solution propriétaire du groupe. Côté tarif, 7 euros par utilisateur et par mois contre 18 pour Kizeo, mais la valeur attendue est surtout la conformité et l’exploitation des données.",(c)=>{
 c.T(128,200,1600,100,'Une décision du groupe, pas un choix local',{f:FH,size:64,bold:1});
 c.T(128,300,1100,40,'Comparatif du cadrage : Kizeo (externe) face à Teepee (solution interne VINCI)',{size:28,color:c.k.mu});
 c.T(940,370,400,44,'KIZEO',{size:28,bold:1,color:c.k.wm,sp:3,align:'center'});c.T(1380,370,412,44,'TEEPEE',{size:28,bold:1,color:c.k.ac,sp:3,align:'center'});
 [['Conformité sécurité VINCI','Non','Oui'],['Tableaux de bord','Non','Oui'],['Automatisation des processus','Oui','Oui'],['Plans et schémas','Non','Oui'],['Connecteur Power BI','Non','Oui'],['Tarif par utilisateur et par mois','18 €','7 €']].forEach((r,i)=>{const y=430+i*88;
  c.R(128,y,1664,2,{fill:c.k.ln});c.T(128,y+10,780,76,r[0],{size:36,valign:'middle'});
  c.T(940,y+10,400,76,r[1],{f:FH,size:44,bold:1,color:r[1]=='Non'?c.k.wm:c.k.tx,align:'center',valign:'middle'});c.T(1380,y+10,412,76,r[2],{f:FH,size:44,bold:1,color:c.k.ac,align:'center',valign:'middle'});});
 c.R(128,958,1664,2,{fill:c.k.ln});
});
def('Problématique',false,[1350,550,900],"Kizeo était maîtrisé par les équipes terrain, alors que Teepee présentait au démarrage des bugs, des lenteurs et des difficultés de synchronisation. La question n’était donc pas seulement technique : comment conduire le changement tout en gardant la continuité opérationnelle ?",(c)=>{
 c.T(128,250,200,260,'“',{f:FH,size:360,color:c.k.ac,bold:1});
 c.T(128,215,500,40,'Problématique',{size:28,bold:1,color:c.k.ac,sp:3,up:1});
 c.T(300,330,1450,470,'Comment conduire une migration numérique attendue par le groupe sans dégrader la performance opérationnelle ni compromettre l’adhésion des équipes terrain, dans un contexte de maturité progressive de la solution cible ?',{f:FH,size:54,lh:1.25});
 c.R(300,860,430,70,{fill:c.k.ac,rad:0.5});c.T(300,860,430,70,'Continuité opérationnelle',{size:28,bold:1,color:'FFFFFF',align:'center',valign:'middle'});
 c.R(760,860,380,70,{fill:c.k.tx,rad:0.5});c.T(760,860,380,70,'Adhésion des équipes',{size:28,bold:1,color:c.k.bg,align:'center',valign:'middle'});
});
def('Pilotage',true,[-300,-300,700],"Il n’y a pas eu de framework agile formel, mais un fonctionnement itératif : tester, recueillir les retours, corriger, redéployer et vérifier. La migration a été vue comme une occasion de rationaliser, pas de recopier l’existant.",(c)=>{
 c.T(128,200,1600,100,'Une démarche itérative en cinq temps',{f:FH,size:64,bold:1});
 c.R(220,520,1480,4,{fill:c.k.ln});
 [['Cartographier','formulaires actifs, obsolètes et doublons'],['Rationaliser','garder l’essentiel, simplifier le reste'],['Paramétrer','formulaires, workflows et modèles d’export'],['Tester','fonctionnel, métier, workflow, export, terrain'],['Former et déployer','janvier 2026, puis stabilisation']].forEach((r,i)=>{const cx=294+i*333;
  c.E(cx-44,478,88,88,{fill:i==4?c.k.ac:c.k.bg,line:c.k.ac,lw:3});c.T(cx-44,478,88,88,String(i+1),{f:FH,size:40,bold:1,color:i==4?c.k.bg:c.k.ac,align:'center',valign:'middle'});
  c.T(cx-150,600,300,100,r[0],{f:FH,size:36,bold:1,align:'center',lh:1.1});c.T(cx-150,700,300,140,r[1],{size:26,color:c.k.mu,align:'center',lh:1.25});});
 c.R(128,880,1664,80,{fill:c.k.pn,line:c.k.ln,lw:2,rad:0.5});
 c.T(128,880,1664,80,'Boucle de travail : tester · recueillir les retours · corriger · redéployer · vérifier',{size:30,align:'center',valign:'middle',color:c.k.ac,bold:1});
});
def('Pilotage',false,[1450,650,800],"Audit, paramétrage, tests, formation et déploiement ont été réalisés. La coupure de Kizeo a été décalée d’un mois, du 31 décembre 2025 au 31 janvier 2026. Ce n’est pas un simple retard : c’est un arbitrage pour réduire le risque de rupture opérationnelle. La stabilisation est devenue une phase à part entière.",(c)=>{
 c.T(128,200,1600,100,'Planning : un mois de plus, par choix',{f:FH,size:64,bold:1});
 const [w,h]=fit('gantt.png',1020,452);c.R(128,320,1060,500,{fill:'FFFFFF',line:c.k.ln,lw:2,rad:0.05});c.I(158,344+(452-h)/2,w,h,'gantt.png','Chronologie synthétique du projet');
 c.T(128,840,1060,40,'Figure 1 : chronologie synthétique du projet, en mois relatifs au lancement',{size:24,italic:1,color:c.k.mu});
 c.T(1260,320,540,150,'+1 mois',{f:FH,size:100,bold:1,color:c.k.ac});
 c.T(1260,470,540,120,'Coupure de Kizeo : du 31/12/2025 au 31/01/2026',{size:34,bold:1,lh:1.2});
 c.T(1260,620,530,220,'Un arbitrage pour sécuriser la transition et réduire le risque de rupture, pas un simple retard. La stabilisation est devenue une phase à part entière.',{size:28,color:c.k.mu,lh:1.3});
});
def('Pilotage',true,[1500,-350,800],"L’instabilité de Teepee était le principal risque : perte de confiance des utilisateurs et comparaison défavorable avec Kizeo. J’ai centralisé les anomalies, je les ai classées par criticité, puis suivies avec les développeurs. Mon rôle a évolué : piloter le problème jusqu’à sa résolution plutôt que le résoudre moi-même.",(c)=>{
 c.T(128,200,1600,100,'Un bug se pilote en cinq niveaux',{f:FH,size:64,bold:1});
 [['Identifier','Que se passe-t-il ?','Le formulaire ne se synchronise pas'],['Qualifier','Dans quelles conditions ?','Réseau faible, formulaire X, Android'],['Prioriser','Quel impact ?','Bloquant pour l’intervention'],['Escalader','Qui doit agir ?','L’équipe méthode, puis Teepee'],['Vérifier','Le correctif fonctionne-t-il ?','Test de non-régression']].forEach((r,i)=>{const y=330+i*124;
  c.E(128,y+14,76,76,{fill:c.k.ac});c.T(128,y+14,76,76,String(i+1),{f:FH,size:36,bold:1,color:c.k.bg,align:'center',valign:'middle'});
  c.T(240,y+14,500,76,r[0],{f:FH,size:44,bold:1,valign:'middle'});c.T(760,y+14,520,76,r[1],{size:32,color:c.k.mu,valign:'middle'});c.T(1290,y+14,502,76,r[2],{size:30,color:c.k.ac,bold:1,valign:'middle'});
  if(i<4)c.R(128,y+110,1664,2,{fill:c.k.ln});});
});
def('Pilotage',false,[-350,600,800],"J’ai utilisé le modèle ADKAR pour conduire le changement. L’étape critique est la lettre A de Ability : la capacité à utiliser l’outil malgré les bugs, grâce à la formation et à un accompagnement après la formation, en janvier 2026, d’abord auprès des équipes terrain.",(c)=>{
 c.T(128,200,1600,100,'Conduire le changement : le modèle ADKAR',{f:FH,size:64,bold:1});
 [['A','Awareness','Expliquer les raisons du groupe et de la sécurité'],['D','Desire','Impliquer les utilisateurs dans les remontées'],['K','Knowledge','Former aux fonctions Teepee et aux formulaires'],['A','Ability','Accompagner l’usage en situation réelle'],['R','Reinforcement','Suivre l’adoption et communiquer les progrès']].forEach((r,i)=>{const x=128+i*340,hot=i==3;
  c.T(x,300,300,240,r[0],{f:FH,size:220,bold:1,color:hot?c.k.ac:'C9C4B8',lh:1});c.R(x,560,hot?300:100,hot?10:4,{fill:hot?c.k.ac:c.k.ln});
  c.T(x,590,300,50,r[1],{f:FH,size:36,bold:1});c.T(x,650,290,160,r[2],{size:28,color:c.k.mu,lh:1.25});});
 c.R(128+3*340,860,300,70,{fill:c.k.ac,rad:0.5});c.T(128+3*340,860,300,70,'Étape critique',{size:28,bold:1,color:'FFFFFF',align:'center',valign:'middle'});
});
def('Pilotage',true,[1400,600,900],"Le cadrage identifiait six risques. Les trois les plus critiques, au niveau 4 sur 4, étaient l’instabilité technique, la perte de confiance et la perte de données. La réponse : suivi des anomalies, formation, accompagnement et maintien temporaire de Kizeo.",(c)=>{
 c.T(128,200,1600,100,'Les risques, du plus au moins critique',{f:FH,size:64,bold:1});
 [['Instabilité technique',4],['Perte de confiance',4],['Perte de données',4],['Retard de planning',3],['Maîtrise de l’outil',2],['Résistance au changement',2]].forEach((r,i)=>{const y=330+i*100,hot=r[1]==4;
  c.T(128,y,520,64,r[0],{size:34,valign:'middle'});c.R(680,y+10,1000,44,{fill:c.k.pn,rad:0.5});c.R(680,y+10,1000*r[1]/4,44,{fill:hot?c.k.wm:c.k.ac,rad:0.5});
  c.T(1710,y,90,64,r[1]+' / 4',{f:FH,size:32,bold:1,valign:'middle'});});
 c.T(128,950,1664,40,'Criticité de 1 (faible) à 4 (critique), évaluation qualitative issue de la note de cadrage',{size:24,color:c.k.mu,italic:1});
});
// showcase : 6 étapes, la bande de vignettes reste, l’image active grandit (Morph)
const SH=[['a1_cartographie.png','Cartographie','Inventaire des formulaires Kizeo : usages, doublons et formulaires obsolètes.','Figure 2 : cartographie des formulaires Kizeo'],
['a5_listes.png','Import des données','Les données sont importées directement dans Teepee : 84 % des formulaires avec données.','Figure 3 : listes de données dans Teepee'],
['a5_formulaire.png','Paramétrage','Chaque formulaire est reconstruit dans l’éditeur de Teepee, avec ses champs et ses modèles d’export.','Figure 4 : éditeur de formulaire Teepee'],
['a5_workflow.png','Workflows','Automatiser la validation, par exemple par l’envoi de mails.','Figure 5 : workflow d’envoi de mails'],
['a2_teepee_mobile.png','Terrain','Saisie sur mobile par les équipes terrain, y compris hors réseau.','Figure 6 : saisie d’un formulaire sur mobile'],
['a2_guide.png','Formation','Supports de formation, janvier 2026, d’abord pour les équipes terrain. Le journal des anomalies a servi à piloter la stabilisation.','Figure 7 : guide de formation Teepee']];
const SHN=["Première étape : la cartographie des formulaires Kizeo, avec leurs usages, pour distinguer les formulaires actifs, obsolètes ou redondants.","Les données ont ensuite été importées directement dans Teepee. 84 % des formulaires avec données ont pu l’être ; pour le reste, les données ont été sauvegardées.","Chaque formulaire a été reconstruit dans l’éditeur de Teepee, avec ses champs, ses composants et ses modèles d’export.","Les workflows automatisent la validation et les notifications, par exemple l’envoi de mails.","Côté terrain, les utilisateurs saisissent depuis l’application mobile, y compris hors réseau.","Enfin les supports de formation, réalisés en janvier 2026, et le journal des anomalies qui a piloté la stabilisation."];
SH.forEach((s,k)=>def('Résultats',false,[1500+k*40,700-k*60,800],SHN[k],(c)=>{
 c.T(128,190,1500,90,'Du formulaire Kizeo au terrain',{f:FH,size:56,bold:1});
 c.R(128,290,1100,470,{fill:'FFFFFF',line:c.k.ln,lw:2,rad:0.04,name:'mcard'});
 SH.forEach((q,j)=>{const x=128+j*254,y=850;c.R(x,y,240,120,{fill:'FFFFFF',line:c.k.ln,lw:2,rad:0.06,name:'tf'+j});});
 c.R(128+k*254-6,844,252,132,{fill:null,line:c.k.ac,lw:5,rad:0.07,name:'sel'});
 SH.forEach((q,j)=>{if(j===k){const [w,h]=fit(q[0],1060,430);c.I(128+(1100-w)/2,290+(470-h)/2,w,h,q[0],q[1],'shot'+j);}else{const [w,h]=fit(q[0],220,100);c.I(128+j*254+(240-w)/2,850+(120-h)/2,w,h,q[0],q[1],'shot'+j);}});
 c.T(128,772,1100,40,s[3],{size:24,italic:1,color:c.k.mu});
 c.T(1300,300,492,40,'Étape '+(k+1)+' sur 6',{size:26,bold:1,color:c.k.ac,sp:3,up:1});
 c.T(1300,350,492,120,s[1],{f:FH,size:56,bold:1,lh:1.05});
 c.T(1300,490,492,300,s[2],{size:30,color:c.k.mu,lh:1.3});
}));
def('Résultats',true,[-300,500,700],"Tous les formulaires actifs ont été migrés. 70 utilisateurs sur 91 ont été formés, soit environ 77 %, ce qui couvre tous les utilisateurs actifs et réguliers. L’import atteint 84 % : les données de certains formulaires n’ont pas pu être importées à cause de différences de fonctionnement entre les deux solutions, mais elles ont été sauvegardées. Les incidents bloquants sont passés sous un par mois.",(c)=>{
 c.T(128,200,1600,100,'Ce que disent les chiffres',{f:FH,size:64,bold:1});
 [['Formulaires actifs migrés',100,'100 %','cible 100 %'],['Utilisateurs formés',77,'≈ 77 %','70 sur 91, tous les actifs'],['Formulaires avec données importés',84,'84 %','le reste a été sauvegardé']].forEach((r,i)=>{const y=330+i*170;
  c.T(128,y,900,50,r[0],{size:36,bold:1});c.T(128,y+54,900,40,r[3],{size:26,color:c.k.mu});
  c.R(1000,y+20,560,40,{fill:c.k.pn,rad:0.5});c.R(1000,y+20,560*r[1]/100,40,{fill:c.k.ac,rad:0.5});c.T(1590,y-8,220,100,r[2],{f:FH,size:60,bold:1,valign:'middle'});});
 c.R(128,850,1664,2,{fill:c.k.ln});
 c.T(128,872,900,50,'Incidents bloquants',{size:36,bold:1});c.T(128,926,900,40,'cible initiale : moins de 1 par semaine',{size:26,color:c.k.mu});
 c.T(1000,862,792,100,'moins de 1 par mois',{f:FH,size:60,bold:1,color:c.k.ac,align:'right',valign:'middle'});
});
def('Bilan',false,[1500,-300,700],"Quatre objectifs sur cinq sont atteints. L’import de données est partiel, à 84 %, à cause de différences de fonctionnement entre les deux solutions : les données concernées ont tout de même été sauvegardées.",(c)=>{
 c.T(128,200,1600,100,'Objectifs initiaux : quatre sur cinq atteints',{f:FH,size:64,bold:1});
 [['Migrer les formulaires actifs','100 % migrés',1],['Arrêter Kizeo','coupure réalisée le 31/01/2026',1],['Former les utilisateurs','≈ 77 % formés, tous les actifs',1],['Importer les données','84 % des formulaires avec données',0],['Stabiliser l’outil','moins de 1 incident bloquant par mois',1]].forEach((r,i)=>{const y=320+i*124;
  c.E(128,y+12,84,84,{fill:r[2]?c.k.ac:'FFFFFF',line:r[2]?c.k.ac:c.k.wm,lw:5});c.T(128,y+12,84,84,r[2]?'✓':'84',{f:FB,size:r[2]?44:30,bold:1,color:r[2]?'FFFFFF':c.k.wm,align:'center',valign:'middle'});
  c.T(250,y+4,820,56,r[0],{f:FH,size:40,bold:1,valign:'middle'});c.T(250,y+58,900,44,r[1],{size:28,color:c.k.mu});
  c.R(1560,y+24,232,60,{fill:r[2]?c.k.ac:c.k.wm,rad:0.5});c.T(1560,y+24,232,60,r[2]?'Atteint':'Partiel',{size:28,bold:1,color:'FFFFFF',align:'center',valign:'middle'});
  if(i<4)c.R(128,y+116,1664,2,{fill:c.k.ln});});
});
def('Bilan',true,[1350,500,900],"Ma posture a évolué : d’abord l’exécution et le paramétrage, puis la coordination et le pilotage entre le terrain, les autres entités et l’équipe Teepee. Mes points forts sont la gestion de projet, l’analyse fonctionnelle, la gestion des risques et l’amélioration continue ; je progresse en conduite du changement et en pilotage par la donnée ; la communication reste à un niveau moyen.",(c)=>{
 c.T(128,200,1600,100,'Compétences développées',{f:FH,size:64,bold:1});
 [['Niveau fort',['Gestion de projet','Analyse fonctionnelle','Gestion des risques','Amélioration continue'],c.k.ac],['En progression',['Conduite du changement','Pilotage par la donnée'],c.k.tx],['Niveau moyen',['Communication'],c.k.mu]].forEach((col,i)=>{const x=128+i*570;
  c.T(x,320,520,50,col[0],{size:28,bold:1,color:col[2],sp:3,up:1});c.R(x,376,520,4,{fill:col[2]});
  col[1].forEach((t,j)=>{c.R(x,410+j*100,520,80,{fill:c.k.pn,rad:0.2});c.T(x+28,410+j*100,470,80,t,{size:32,valign:'middle'});});});
 c.T(128,860,1664,90,'Posture : de l’exécution et du paramétrage vers la coordination et le pilotage',{f:FH,size:40,italic:1,color:c.k.ac,valign:'middle'});
});
def('Bilan',false,[-300,-250,700],"La difficulté la plus formatrice a été de travailler avec une solution qui n’était pas totalement stabilisée. Cela m’a appris à distinguer trois niveaux : le problème technique, son impact métier et sa gestion projet. Je ne pouvais pas toujours corriger le bug, mais je pouvais améliorer son escalade, sa priorisation et sa communication.",(c)=>{
 c.T(128,200,1650,100,'Gérer l’incertitude technique',{f:FH,size:64,bold:1});
 [['Le problème technique','Un bug, une lenteur, une synchronisation qui échoue'],['Son impact métier','Productivité, confiance, adhésion des utilisateurs'],['Sa gestion projet','Escalade, priorisation et communication']].forEach((r,i)=>{const x=128+i*570;
  c.E(x,340,120,120,{fill:i==2?c.k.ac:'FFFFFF',line:c.k.ac,lw:4});c.T(x,340,120,120,String(i+1),{f:FH,size:56,bold:1,color:i==2?'FFFFFF':c.k.ac,align:'center',valign:'middle'});
  c.T(x,490,520,110,r[0],{f:FH,size:42,bold:1,lh:1.1});c.T(x,610,500,110,r[1],{size:28,color:c.k.mu,lh:1.3});});
 c.R(128,790,1664,2,{fill:c.k.ln});
 c.T(128,820,1664,120,'Je ne pouvais pas toujours corriger le bug, mais je pouvais améliorer son escalade, sa priorisation et sa communication.',{f:FH,size:40,italic:1,lh:1.25});
});
def('Bilan',true,[1400,600,900],"Avec le recul, trois améliorations auraient renforcé le projet : classer les anomalies dès les premières remontées, mettre en place un planning plus détaillé, et intégrer la stabilisation comme une vraie phase projet avec des ressources et des indicateurs dédiés. Pour un futur projet, j’évaluerais plus tôt la capacité de la solution à supporter le déploiement.",(c)=>{
 c.T(128,200,1600,100,'Avec le recul, trois choses à refaire',{f:FH,size:64,bold:1});
 [['Classer les anomalies','dès les premières remontées, de façon standardisée'],['Planifier plus finement','un planning plus détaillé dès le départ'],['Prévoir la stabilisation','comme une phase projet, avec ressources et indicateurs']].forEach((r,i)=>{const y=330+i*210;
  c.T(128,y-30,220,200,String(i+1),{f:FH,size:200,bold:1,color:c.k.ac,lh:1});c.T(380,y+20,1400,70,r[0],{f:FH,size:54,bold:1});c.T(380,y+95,1400,50,r[1],{size:32,color:c.k.mu});c.R(380,y+170,1412,2,{fill:c.k.ln});});
});
def('Bilan',false,[1500,-300,800],"La prochaine étape est de consolider l’usage de Teepee avant d’élargir son périmètre. Une fois la stabilité suffisante, les données collectées pourront être valorisées : tableaux de bord, Power BI, automatisation de processus et analyse de performance.",(c)=>{
 c.T(128,200,1600,100,'Perspectives : consolider, puis valoriser',{f:FH,size:64,bold:1});
 c.R(128,330,780,560,{fill:'FFFFFF',line:c.k.ln,lw:2,rad:0.05});c.R(1012,330,780,560,{fill:c.k.ac,rad:0.05});
 c.T(172,362,700,50,'D’abord : consolider',{size:28,bold:1,color:c.k.ac,sp:3,up:1});
 ['Registre d’anomalies : criticité, responsable, statut','Suivi mensuel de l’adoption et des incidents','Bonnes pratiques de création des formulaires','Retours des autres entités via la communauté inter-agences'].forEach((t,i)=>{c.R(172,432+i*110,12,12,{fill:c.k.ac});c.T(204,414+i*110,660,100,t,{size:30,lh:1.2});});
 c.T(1056,362,700,50,'Ensuite : valoriser la donnée',{size:28,bold:1,color:'FFFFFF',sp:3,up:1});
 ['Tableaux de bord','Power BI','Automatisation de processus','Analyse de performance'].forEach((t,i)=>{c.R(1056,432+i*110,12,12,{fill:'FFFFFF'});c.T(1088,414+i*110,660,100,t,{size:34,color:'FFFFFF',bold:1,lh:1.2});});
 c.T(908,560,104,100,'→',{f:FB,size:72,bold:1,color:c.k.ac,align:'center',valign:'middle'});
});
def('Conclusion',true,[1000,300,1300],"Pour conclure : la migration, la formation et le déploiement ont été réalisés et Kizeo a été coupé le 31 janvier 2026. Le principal enseignement est la relation entre technique et humain : un bug n’est pas seulement un problème informatique, il touche la confiance et l’adhésion. Personnellement, je suis passé d’exécutant à co-pilote du projet. Merci de votre attention, je suis à votre disposition pour vos questions.",(c)=>{
 c.T(128,260,1300,40,'Conclusion',{size:28,bold:1,color:c.k.ac,sp:3,up:1});
 c.T(128,330,1450,440,'Un déploiement technique ne devient une appropriation durable que si l’on pilote aussi l’humain.',{f:FH,size:76,bold:1,lh:1.18});
 c.T(128,830,700,110,'Merci.',{f:FH,size:96,bold:1,color:c.k.ac});
 c.T(700,858,900,60,'Place à vos questions',{size:36,color:c.k.mu});
});

// ---------- rendu ----------
const N=S.length;
(async()=>{
 const pres=new pptxgen();pres.layout='LAYOUT_WIDE';pres.title='Soutenance Kizeo vers Teepee';
 S.forEach((s,i)=>{const k=s.dark?D:L;
  const lab=`${String(i+1).padStart(2,'0')}  ·  ${s.label}`;
  pres.defineSlideMaster({title:'K_'+i,objects:[{text:{text:lab.toUpperCase(),options:{x:inch(128),y:inch(76),w:inch(900),h:inch(40),fontFace:FB,fontSize:12,bold:true,color:k.mu,charSpacing:3,margin:0,valign:'middle'}}}]});
  const sl=pres.addSlide({masterName:'K_'+i});sl.background={color:k.bg};
  const opt=(q)=>q.name?{objectName:'!!'+q.name}:{};
  const c={k,
   R:(x,y,w,h,q={})=>sl.addShape(q.rad?pres.shapes.ROUNDED_RECTANGLE:pres.shapes.RECTANGLE,Object.assign({x:inch(x),y:inch(y),w:inch(w),h:inch(h),fill:q.fill?{color:q.fill,transparency:q.tr||0}:{type:'none'},line:q.line?{color:q.line,width:(q.lw||2)/2}:{type:'none'}},q.rad?{rectRadius:inch(Math.min(14,Math.min(w,h)*0.2))}:{},opt(q))),
   E:(x,y,w,h,q={})=>sl.addShape(pres.shapes.OVAL,Object.assign({x:inch(x),y:inch(y),w:inch(w),h:inch(h),fill:q.fill?{color:q.fill,transparency:q.tr||0}:{type:'none'},line:q.line?{color:q.line,width:(q.lw||2)/2}:{type:'none'}},opt(q))),
   T:(x,y,w,h,t,q={})=>sl.addText(q.up?t.toUpperCase():t,Object.assign({x:inch(x),y:inch(y),w:inch(w),h:inch(h),fontFace:q.f||FB,fontSize:pt(q.size||32),bold:!!q.bold,italic:!!q.italic,color:q.color||k.tx,align:q.align||'left',valign:q.valign||'top',margin:0,isTextBox:true,fit:'none',lineSpacingMultiple:(q.lh||1.2)/1.2},q.sp?{charSpacing:q.sp}:{},opt(q))),
   I:(x,y,w,h,f,alt,name)=>sl.addImage(Object.assign({path:path.join(AS,f),x:inch(x),y:inch(y),w:inch(w),h:inch(h),altText:alt||''},name?{objectName:'!!'+name}:{})),
   V:(x,y,w,h)=>sl.addMedia({type:'video',path:path.join(AS,'demo.mp4'),cover:'data:image/png;base64,'+fs.readFileSync(path.join(AS,'poster.png')).toString('base64'),x:inch(x),y:inch(y),w:inch(w),h:inch(h)})};
  // décor persistant (animé par Morph) : cercle, plaques logos, barre de progression
  sl.addShape(pres.shapes.ISOSCELES_TRIANGLE,{x:inch(1380+40*Math.sin(i)),y:inch(330+30*Math.cos(i)),w:inch(i===0?1000:900+(i%3)*20),h:inch(i===0?900:800+(i%3)*20),rotate:(i%5)*3-6,fill:{color:k.ac,transparency:s.dark?88:90},line:{type:'none'},objectName:'!!orb'});
  const big=i===0;
  const pv=big?[128,80,400,130,340,113]:[1416,56,230,78,190,63], pi=big?[548,80,220,130,140,83]:[1662,56,130,78,80,47];
  c.R(pv[0],pv[1],pv[2],pv[3],{fill:'FFFFFF',rad:0.2,name:'pv'});c.I(pv[0]+(pv[2]-pv[4])/2,pv[1]+(pv[3]-pv[5])/2,pv[4],pv[5],'vinci.png','Logo VINCI Facilities','lv');
  c.R(pi[0],pi[1],pi[2],pi[3],{fill:'FFFFFF',rad:0.2,name:'pi'});c.I(pi[0]+(pi[2]-pi[4])/2,pi[1]+(pi[3]-pi[5])/2,pi[4],pi[5],'icam.png','Logo Icam','li');
  const bw=Math.max(24,1664*i/(N-1));
  c.R(128,1046,1664,8,{fill:k.ln,rad:0.5,name:'track'});c.R(128,1046,bw,8,{fill:k.ac,rad:0.5,name:'fill'});c.E(128+bw-14,1040,20,20,{fill:k.ac,name:'dot'});
  s.fn(c);sl.addNotes(s.notes);
 });
 const raw=path.join(OUT,'_raw.pptx');await pres.writeFile({fileName:raw});
 const MORPH='<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1200"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
 const zip=await JSZip.loadAsync(fs.readFileSync(raw));
 for(const f of Object.keys(zip.files)){const m=f.match(/^ppt\/slides\/slide(\d+)\.xml$/);if(!m)continue;let x=await zip.file(f).async('string');
  if(+m[1]>=2){if(!x.includes('</p:clrMapOvr>'))throw new Error('clrMapOvr '+f);x=x.replace('</p:clrMapOvr>','</p:clrMapOvr>'+MORPH);}
  x=x.replace(/(<p:cNvPr id="\d+" name=")([^"!]*)(")/g,(mm,a,n,b)=>a+n+'-s'+m[1]+b);
  zip.file(f,x);}
 const dst=path.join(OUT,'Presentation_Kizeo_Teepee_v2.pptx');fs.writeFileSync(dst,await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));fs.unlinkSync(raw);console.log(dst,N);
})();
