// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//ngrok http 3978 --host-header=localhost:3978

const { QnAMaker } = require('botbuilder-ai');
const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes,
    AttachmentLayoutTypes,
    TeamsInfo,
    ActivityTypes
} = require('botbuilder');
const axios = require('axios');

//init global vars
var categorie = "";
var res;
var noTickets = false;
var userProfile;
var userProblem = "Problème ";
var newTicketID = "";
// QnA maker get answer options
const qnaOptions = {
    "top": 10,
    "strictFilters": [{ "name": "categorie", "value": "" }]
};

class BotActivityHandler extends TeamsActivityHandler {
    constructor(configuration) {
        super();
        this.onMessage(async (context, next) => {
            //TurnContext.removeRecipientMention(context.activity);
            console.log(context.activity);

            await this.VerifyUserIsClient(context, configuration);


            await next();
        });

        this.onMembersAdded(async (context, next) => {

            await context.sendActivity("Bonjour !")
            await context.sendActivity('Je suis PROBOT un assistant virtuel, je peux vous aidez avec les actions comme indiqué ci-dessous, si vous souhaitez d\'afficher le menu , écrivez simplement "Menu"');
            await this.ShowMenuCard(context);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    //#################### CHATBOT FUNCTIONS ######################//

    async VerifyUserIsClient(context, configuration) {
        var userTenantID = context.activity.conversation.tenantId;

        const OpenContactPage = CardFactory.adaptiveCard({
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.2",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "Achetez un abonnement !",
                    "fontType": "Default",
                    "size": "Medium",
                    "weight": "Bolder"
                }
            ],
            "actions": [
                {
                    "type": "Action.OpenUrl",
                    "title": "Contacter PROGED",
                    "url": "https://www.progedsolutions.com/contact/",
                    "iconUrl": "https://www.progedsolutions.com/wp-content/uploads/2019/07/proged-cargo.png"
                }
            ]
        });
        await axios
            .post('https://prod-53.westeurope.logic.azure.com:443/workflows/ce20abd21c90478391d62bc0f6ae13f1/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=k5TASE7_9peuCq443tqwv5TWZcmGlwVqPeyOCDanTc4', {
                Tenantid: userTenantID
            })
            .then(async res => {
                if (res.data.existe == "exist0") {////change =="exist" to make all users can access the chatbot functions */
                    await context.sendActivity("Vous ne pouvez pas y accéder pour s'inscrire vous pouvez nous contactez :");
                    await context.sendActivity({ attachments: [OpenContactPage] });
                }
                else {
                    await this.HandleUserMessageRecived(context, configuration);
                }
            })
            .catch(error => {
                console.error(error)
            });

    }

    async HandleUserMessageRecived(context, configuration) {
        
        userProfile = await TeamsInfo.getMembers(context);
        //check if the message recived is from the Menu buttons 
        switch (context.activity.text) {
            case 'problemsCard':
                {
                    const problemCard = CardFactory.heroCard(
                        'Choisissez votre type de problème',
                        null, [{
                            type: ActionTypes.MessageBack,
                            title: "1.1 Problème  Office 365",
                            value: { cardHint: "button", problem: "Problème  Office 365" },
                            text: 'office'
                        }, {
                            type: ActionTypes.MessageBack,
                            title: "1.2 Problème  SharePoint",
                            value: { cardHint: "button", problem: "Problème  SharePoint" },
                            text: 'sharepoint'
                        }, {
                            type: ActionTypes.MessageBack,
                            title: "1.3 Problème Licence",
                            value: { cardHint: "button", problem: "Problème  Licence" },
                            text: 'licence'
                        }]);
                    await context.sendActivity({ attachments: [problemCard] });

                }
                break;

            case 'Get My Tickets':
                context.sendActivity("Attendez, je suis en train de traiter votre demande")
                await this.getAllTicketPerUser(context);
                break;

            case ('menu' || 'Menu'):
                await this.ShowMenuCard(context);
                break;

            case "supportTicket":
                await this.createSupportTicket(context);
                break;

            case 'my infos':
                await this.getUserInfoFromTeams(context);
                break;

            case 'problem_resolu':
                await context.sendActivity('Parfait 😉 , si vous avez besoin d’autre chose tapez "Menu" à tout moment.');
                break;

            case 'NoaddAttachment':
                await context.sendActivity("Tu peux consulter l'état de votre demande d'assistance à tout moment ✅.");
                break;
            //if the message is not from Menu button 
            default:
                //check if the message if from adaptive card
                if (context.activity.value != undefined) {
                    //check if the message is from the button "Poser une question"
                    if (context.activity.value.hasOwnProperty('cardHint')) {
                        await this.getQuestionsFromQna(context, configuration);
                        console.log('got answers from that button');
                    }

                    //check if the message is from add cmnt submit button 
                    else if (context.activity.value.hasOwnProperty('addCmnt')) {
                        var cmnt = await context.activity.value.addCmnt;
                        var tikID = await context.activity.value.ticketID;
                        await this.sendTicketNoteToMantis(context, cmnt, tikID);
                        console.log("######NOTE Function");
                    }
                    else if (context.activity.value.hasOwnProperty('cardType')) {
                        await this.getAnswerFromQna(context, configuration);
                        console.log("////////////////////////////////////////////////////");
                    }
                    else if (context.activity.value.hasOwnProperty('ticketIDrecived')) {
                        newTicketID = context.activity.value.ticketIDrecived.toString();
                        await context.sendActivity("veuillez ajouter la pièce jointe ici ⬇⬇");
                    }
                    else {
                        //check the message if from ticket adaptive card submit button
                        if (context.activity.value.hasOwnProperty('ticketProblem')) {
                            await this.sendTicketToApiMantis(context);
                            console.log("***************create ticket function");
                        }
                    }
                }
                else if (context.activity.hasOwnProperty('attachments')) {
                    await this.sendTicketAttachmentToMantis(context);

                }
                //pass the message to getQuestionsFromQna to check if there is answer or return the Menu again
                else {
                    if (!configuration) throw new Error('[QnaMakerBot]: Missing parameter. configuration is required');
                    await this.getQuestionsFromQna(context, configuration);
                    console.log('the default switch statment');
                }

                break;
        }
    }

    async getAllTicketPerUser(context) {
        let nameuser = context.activity.from.name;

        var ticketsArray = [];
        await axios
            .post('https://prod-191.westeurope.logic.azure.com:443/workflows/dbd1955c245b4e9982b473b37b62e464/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sp6o4IdUWWu5k9_T7QXnJAkNAk3UWT8oVkr1KsY0iiU', {
                Reporter: userProfile[0].email
            })
            .then(async function (res) {
                if (res.data.issues.length == 0) {
                    context.sendActivity("Vous n'avez pas des demandes d'assistance");
                    noTickets = true;
                }
                //console.log(`statusCode: ${res.statusCode}`)
                if (res.data.issues.length != 0) {
                    for (let i = 0; i < res.data.issues.length; i++) {
                        if (res.data.issues[i].id.toString() == "715") {
                            res.data.issues[i].description = "une longue description";
                        }
                        //await context.sendActivity("Ticket ID : " + res.data.issues[i].id.toString() + " <br/>  Sujet     : " + res.data.issues[i].summary.toString());
                        let cardd = CardFactory.adaptiveCard({
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "text": `Ticket ID : ${res.data.issues[i].id.toString()}`,
                                    "wrap": true,
                                    "id": "ticketId",
                                    "horizontalAlignment": "Center"
                                },
                                {
                                    "type": "ColumnSet",
                                    "columns": [
                                        {
                                            "type": "Column",
                                            "items": [
                                                {
                                                    "type": "Image",
                                                    "style": "Person",
                                                    "url": "https://cdn4.iconfinder.com/data/icons/business-charts-rounded/512/xxx022-512.png",
                                                    "size": "Small"
                                                }
                                            ],
                                            "width": "auto"
                                        },
                                        {
                                            "type": "Column",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "weight": "Bolder",
                                                    "wrap": true,
                                                    "text": `${nameuser}`,
                                                    "id": "userName"
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "spacing": "None",
                                                    "text": `Date de création : ${res.data.issues[i].created_at.toString()}`,
                                                    "isSubtle": true,
                                                    "wrap": true,
                                                    "id": "ticketDate"
                                                }
                                            ],
                                            "width": "stretch"
                                        }
                                    ]
                                },
                                {
                                    "type": "TextBlock",
                                    "size": "Medium",
                                    "weight": "Bolder",
                                    "text": `Objet de demande d'assistance : ${res.data.issues[i].summary.toString()}`,
                                    "height": "stretch",
                                    "color": "Good"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Descritption :",
                                    "wrap": true,
                                    "id": "ticketDesctitle"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": `${res.data.issues[i].description.toString()}`,
                                    "wrap": true,
                                    "id": "ticketDesc"
                                }
                            ],
                            "actions": [
                                {
                                    "type": "Action.ShowCard",
                                    "title": "Ajouter commentaire",
                                    "card": {
                                        "type": "AdaptiveCard",
                                        "body": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Ajouter votre commentaire",
                                                "wrap": true,
                                                "weight": "Bolder"
                                            },
                                            {
                                                "type": "Input.Text",
                                                "placeholder": "Votre commentaire...",
                                                "isMultiline": true,
                                                "id": "addCmnt"
                                            }
                                        ],
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "Envoyer",
                                                "style": "positive",
                                                "data": {
                                                    "ticketID": `${res.data.issues[i].id}`

                                                }
                                            }
                                        ],
                                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"
                                    },
                                    "style": "positive"
                                }
                            ],
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.3"
                        }
                        );
                        ticketsArray.push(cardd)
                    }
                }

            })
            .catch(error => {
                console.error(error)
            });

        if (noTickets == false) {
            await context.sendActivity({
                attachments: ticketsArray,
                attachmentLayout: AttachmentLayoutTypes.Carousel
            });
        }
    }

    async getUserInfoFromTeams(context) {
        let nameFromTeams = context.activity.from.name;
        let aadIdFromTeams = context.activity.from.aadObjectId;
        let userIdFromTeams = context.activity.from.id;


        //let tenantIdFromTeams = context.activity.from.conversation.tenantId;
        await context.sendActivity(`Votre nom : ${userProfile[0].name}`);
        await context.sendActivity(`Votre Azure AD ID : ${userProfile[0].tenantId}`);
        await context.sendActivity(`Votre Teams user ID : ${userProfile[0].id}`);
    }

    async createSupportTicket(context) {
        //get required data for ticket adaptive card
        let options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        let currentDate = new Date().toLocaleDateString("fr-FR", options);
        let nameFromTeams = context.activity.from.name;
        await context.sendActivity("Veuillez remplir cet demande d'assistance pour résoudre votre problème avec l'un de nos agents de support :")
        //create dynamic adaptive card for support ticket
        let SupportTicket = CardFactory.adaptiveCard({
            "type": "AdaptiveCard",
            "body": [{
                "type": "TextBlock",
                "text": "Création d'une demande d'assistance :",
                "wrap": true
            },
            {
                "type": "ColumnSet",
                "columns": [{
                    "type": "Column",
                    "items": [{
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": nameFromTeams,
                        "wrap": true
                    },

                    {
                        "type": "TextBlock",
                        "spacing": "None",
                        "text": currentDate,
                        "isSubtle": true,
                        "wrap": true
                    }
                    ],
                    "width": "stretch"
                }]
            },
            {
                "type": "Input.Text",
                "text": "dsdsd",
                "placeholder": `${userProblem}`,
                "id": "ticketProblem"
            },
            {
                "type": "Input.Text",
                "placeholder": "Description",
                "isMultiline": true,
                "id": "ticketDescription"
                // ,"inlineAction": {
                //     "type": "Action.Submit",
                //     "id": "attachmentBtn",
                //     "title": "Ajouter capture",
                //     "data": {
                //         "attachment": true
                //     }
                // }
            },
            {
                "type": "Input.ChoiceSet",
                "choices": [{
                    "title": "Low ",
                    "value": "Low "
                }, {
                    "title": "immédiate ",
                    "value": "immédiate "
                },
                {
                    "title": "Normale",
                    "value": "Normale"
                }
                ],
                "placeholder": "Priorité",
                "id": "ticketType"
            }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2",
            "actions": [{
                "type": "Action.Submit",
                "title": "Envoyer",
                "id": "submitButton"
            }]
        });
        // SupportTicket.content.body[1].columns[0].items[0].text = nameFromTeams;
        // SupportTicket.content.body[1].columns[0].items[1].text = currentDate;
        await context.sendActivity({ attachments: [SupportTicket] });

    }

    async sendTicketToApiMantis(context) {
        //validate if the auto problem detection is true else get the user mannuly inserted value
        var summary = '';
        summary = context.activity.value.ticketProblem;
        if (summary == '') {
            summary = userProblem;
        }
        let description = context.activity.value.ticketDescription;
        let nameFromTeams = context.activity.from.name;
        let tenantId = context.activity.conversation.tenantId;
        let Reporter = nameFromTeams;
        let Priority = context.activity.value.ticketType;
        if (description != "" && summary != "") {
            await axios
                .post('https://prod-25.westeurope.logic.azure.com:443/workflows/324f76aa3aa34fc0bd3c5700e96586a5/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=htXVckB5lWbyAher2ux9lRhWSy8opoupEIsPyeQXRDY', {
                    summary: summary,
                    description: description,
                    Reporter: userProfile[0].email,
                    Priority: Priority,
                    project: { TenantId: tenantId, name: '' }
                })
                .then(async res => {
                    newTicketID = res.data.issue.id;
                    await context.sendActivity(`La demande d'assistance a été bien enregistrée ✅ <br/> ID de la demande : ${res.data.issue.id} <br/> Vous recevez un E-mail d'enregistrement de votre demande .`);
                    await context.sendActivity("Tu peux consulter l'état de votre demande d'assistance à tout moment.");
                })
                .catch(error => {
                    console.error(error)
                });

            await this.askForTicketAttachment(context);
        }
        else {
            await context.sendActivity("Veuillez vérifier que vous avez rempli tous les champs de demande d'assistance ci-dessus 🧐");
        }

    }

    async askForTicketAttachment(context) {
        const addAttachmentCard = CardFactory.heroCard(
            "Voulez vous ajoutez un attachement sur cet demande d'assistance !",
            null, [{
                type: ActionTypes.MessageBack,
                title: "Oui",
                text: 'addAttachment',
                value: { "ticketIDrecived": newTicketID }
            }, {
                type: ActionTypes.MessageBack,
                title: "Non",
                text: 'NoaddAttachment',
                value: {}
            }]);
        await context.sendActivity({ attachments: [addAttachmentCard] });
    }

    async sendTicketAttachmentToMantis(context) {
        var attachmentContent = context.activity.attachments[0].content.downloadUrl.toString();
        var attachmentName = context.activity.attachments[0].name.toString();
        await axios
            .post('https://prod-25.westeurope.logic.azure.com:443/workflows/324f76aa3aa34fc0bd3c5700e96586a5/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=htXVckB5lWbyAher2ux9lRhWSy8opoupEIsPyeQXRDY', {
                addedAttachment: "true",
                IssueId: parseInt(newTicketID),
                attachmentContent: attachmentContent,
                attachmentName: attachmentName,
                summary: '...',
                description: '...',
                Reporter: '...',
                Priority: '...',
                project: {}
            })
            .then(res => {
                context.sendActivity("Votre Attachement a été bien ajouté ✅");
                newTicketID = "";
            })
            .catch(error => {
                console.error(error)
            });

    }

    async sendTicketNoteToMantis(context, cmnt, ticketID) {
        var id = ticketID;
        await axios
            .post('https://prod-181.westeurope.logic.azure.com:443/workflows/bbaf2b58acf742e4a15338686b8e6b7a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xKnSKgkbMCYSyWDMzYSbz14kWJs8SrWlVfRiOp0jlDs', {
                IssueId: id,
                ReporterEmail: userProfile[0].email,
                text: cmnt,
                files: [{ name: '', content: '' }]
            })
            .then(res => {
            })
            .catch(error => {
                console.error(error)
            });
        context.sendActivity("Votre commentaire a été bien ajouté ✅ nous allons vous répondre le plus tôt possible ");

    }

    /*Pass message to QnA maker to get response */
    async getQuestionsFromQna(context, configuration) {

        //set qna filter to get user requested catergorie of questions
        qnaOptions.strictFilters[0].value = context.activity.text;
        //console.log(qnaOptions);

        this.qnaMaker = new QnAMaker(configuration, qnaOptions);
        const qnaResults = await this.qnaMaker.getAnswers(context, qnaOptions);
        //console.log(qnaResults);

        //hero card that contains qna maker answers for the selected topic

        let values = { cardType: "questions" };

        let questionCardArray = [];

        // If an answer was received from QnA Maker, send the answer back to the user.
        if (qnaResults[0]) {
            //set the var userProblem to the user problem type from adaptive card
            userProblem = await context.activity.value.problem;
            for (let i = 0; i < qnaResults.length; i++) {
                let btn = { type: ActionTypes.MessageBack, title: `● ${qnaResults[i].questions[0]}`, value: values, text: ` ${qnaResults[i].questions[0].toString()}` }
                questionCardArray.push(btn);
            }
            questionCardArray.push({ type: ActionTypes.MessageBack, title: `● autre`, value: values, text: "supportTicket" });
            const questionsCard = CardFactory.heroCard(
                "Choisissez l'une des questions",
                null, questionCardArray);
            await context.sendActivity({ attachments: [questionsCard] });
            console.log(questionsCard.content.buttons[0]);
        }
        // If no answers were returned from QnA Maker, reply the Menu adaptive card.
        else {
            await context.sendActivity('Veuillez sélectionnez un bouton du menu :');
            await this.ShowMenuCard(context);
        }
    }

    //check this one problem in qnaoptions
    async getAnswerFromQna(context, configuration) {
        //set qna filter to get user requested catergorie of questions
        this.qnaMaker = new QnAMaker(configuration);

        const qnaResults = await this.qnaMaker.getAnswers(context);

        // If an answer was received from QnA Maker, send the answer back to the user.
        if (qnaResults[0]) {
            await context.sendActivity(qnaResults[0].answer);
            await this.verifyProblemSolved(context);
            // If no answers were returned from QnA Maker, reply with help.
        } else {
            await context.sendActivity('No answers were found please select Button from adaptive card.');
        }
        // await context.sendActivity("Si vous venez d’acheter Office et que vous avez un package physique, la première chose à faire est d’utiliser la clé de produit incluse dans le paquet. <br/> <br/> Ou, si vous avez acheté Office auprès d’un revendeur en ligne, vous avez peut-être reçu la clé de produit par courrier électronique ou elle est écrite sur votre facture. <br/> <br/> Pour utiliser cette clé, accédez à la page setup.office.com, connectez-vous avec un compte Microsoft existant ou créez-en un. <br/> <br/> Vous ne devez utiliser cette clé qu’une seule fois.");
    }

    async verifyProblemSolved(context) {

        const addAttachmentCard = CardFactory.heroCard(
            'votre problème est-il résolu?',
            null, [{
                type: ActionTypes.MessageBack,
                title: "Oui",
                text: 'problem_resolu',
            }, {
                type: ActionTypes.MessageBack,
                title: "Non",
                text: 'supportTicket'
            }]);
        await context.sendActivity({ attachments: [addAttachmentCard] });

    }

    async ShowMenuCard(context) {
        // By default for unknown activity sent by user show
        // a card with the available actions.
        const value = { cardHint: "button" };
        const card = CardFactory.heroCard(
            'Choisissez une commande',
            null, [{
                type: ActionTypes.MessageBack,
                title: "1. Posez une question",
                value: value,
                text: 'problemsCard',
                //displayText: '1'
            },
            {
                type: ActionTypes.MessageBack,
                title: "2. Demandez assistance",
                value: value,
                text: 'supportTicket',
                //displayText: '2'
            },
            {
                type: ActionTypes.MessageBack,
                title: "3. Mes demandes d'assistances",
                value: value,
                text: 'Get My Tickets',
                //displayText: '3'
            }
            // ,
            // {
            //     type: ActionTypes.MessageBack,
            //     title: "4. Mes informations d'identification",
            //     value: value,
            //     text: 'my infos',
            //     displayText: '4'
            // }
        ]);
        await context.sendActivity({ attachments: [card] });
    }

    //Say hello and @ mention the current user.
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi user : ${mention.text}`);
        replyActivity.entities = [mention];

        await context.sendActivity(replyActivity);
    }
}


module.exports.BotActivityHandler = BotActivityHandler;