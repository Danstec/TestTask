using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace EmailSender
{
    public class Program : CodeActivity
    {
        #region Объявление входящих и выходящих аргументов
        [Input("TargetEmail")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> TargetEmail { get; set; }
        #endregion

        /// <summary>
        /// Добавляет в электронное письмо вложения из примечаний контакта
        /// </summary>
        /// <param name="service">Сервис для получения метаданных</param>
        /// <param name="ContactEntityId">Id записи контакта</param>
        /// <param name="TargetEmailRef">Ссылка на сущность электронного письма</param>
        /// <param name="isAttachmentExist">флаг наличия примечаний у контакта</param>
        /// <returns>Возвращает сущность электронного письма с вложениями</returns>
        private Entity AddAttachmentToEmail(IOrganizationService service, Guid ContactEntityId, EntityReference TargetEmailRef, ref bool isAttachmentExist)
        {
            Entity resultEmailEntity = service.Retrieve("email", TargetEmailRef.Id, new ColumnSet(true));

            //throw new Exception(resultEmailEntity["emailid"].ToString());

            //Построение запроса для получения примечаний...
            ConditionExpression noteAttachmentsObjectIdCondition = new ConditionExpression()
            {
                AttributeName = "objectid",
                Operator = ConditionOperator.Equal,
                Values = { ContactEntityId }
            };

            FilterExpression noteAttachmentsQueryFilter = new FilterExpression()
            {
                Conditions = { noteAttachmentsObjectIdCondition },
                FilterOperator = LogicalOperator.And
            };

            QueryExpression noteAttachmentsQuery = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet(new string[] { "subject", "mimetype", "filename", "documentbody" }),
                Criteria = noteAttachmentsQueryFilter
            };

            //Получение коллекции примечаний
            EntityCollection attachmentCollection = service.RetrieveMultiple(noteAttachmentsQuery);

            //Прикрепление полученных примечаний к электронном письму как вложений
            if (attachmentCollection.Entities.Count > 0)
            {
                //Установка флага наличия примечаний
                isAttachmentExist = true;

                EntityCollection mimeCollection = new EntityCollection();

                for (int i = 0; i < attachmentCollection.Entities.Count; i++)
                {
                    mimeCollection.Entities.Add(new Entity("activitymimeattachment"));
                    
                    if (attachmentCollection.Entities[i].Contains("subject"))
                        mimeCollection.Entities[i]["subject"] = attachmentCollection.Entities[i].GetAttributeValue<string>("subject");

                    if (attachmentCollection.Entities[i].Contains("filename"))
                        mimeCollection.Entities[i]["filename"] = attachmentCollection.Entities[i].GetAttributeValue<string>("filename");

                    if (attachmentCollection.Entities[i].Contains("documentbody"))
                        mimeCollection.Entities[i]["body"] = attachmentCollection.Entities[i].GetAttributeValue<string>("documentbody");

                    if (attachmentCollection.Entities[i].Contains("mimetype"))
                        mimeCollection.Entities[i]["mimetype"] = attachmentCollection.Entities[i].GetAttributeValue<string>("mimetype");

                    mimeCollection.Entities[i]["objectid"] = new EntityReference("email", resultEmailEntity.Id);
                    mimeCollection.Entities[i]["objecttypecode"] = "email";

                    service.Create(mimeCollection.Entities[i]);
                }
            }

            return resultEmailEntity;
        }

        protected override void Execute(CodeActivityContext context)
        {
            #region Получение контекста и сервиса
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationService service = context.GetExtension<IOrganizationServiceFactory>().CreateOrganizationService(workflowContext.UserId);
            #endregion

            //Флаг наличия примечаний у контакта
            bool isAttachmentExist = false;

            //Получение контакта, на котором вызван процесс
            Entity ContactEntity = (Entity)service.Retrieve("contact", workflowContext.PrimaryEntityId, new ColumnSet(new string[] { "contactid" }));

            //Добавление вложений в электронное письмо
            Entity email = AddAttachmentToEmail(service, ContactEntity.Id, TargetEmail.Get(context), ref isAttachmentExist);

            #region Отправка электронного письма
            SendEmailRequest sendEmail = new SendEmailRequest();
                sendEmail.EmailId = email.Id;
                //sendEmail.TrackingToken = "";
                sendEmail.IssueSend = true;
                SendEmailResponse res = (SendEmailResponse)service.Execute(sendEmail);
            #endregion

            //Получение Id и Reference пользователя, который инициировал процесс
            Guid initiatorId = workflowContext.InitiatingUserId;
            EntityReference initiatiorRef = new EntityReference("systemuser", initiatorId);

            //Связывание отправленного письма с вложениями и пользователя
            if (isAttachmentExist)
            {
                email["stec_user"] = initiatiorRef;
                service.Update(email);
            }

            #region Обновление количества отправленных писем пользователем
            Entity initiator = service.Retrieve("systemuser", initiatorId, new ColumnSet(new string[] { "stec_send_emails_count" }));

            if (initiator["stec_send_emails_count"] != null)
            {
                var count = initiator["stec_send_emails_count"];
                initiator["stec_send_emails_count"] = (int)count + 1;
            }
            else initiator["stec_send_emails_count"] = 1;

            service.Update(initiator);
            #endregion
        }
    }
}
