using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace SPS2015.RemoteEventReceiversFinalWeb
{
    public class RemoteEventReceiverManager
    {
        private const string RECEIVER_NAME = "ItemAddedEvent";
        private const string LIST_TITLE = "Global Navigation";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            //Get the list collection
            ListCollection listCollection = clientContext.Web.Lists;

            //Load the list Collection
            clientContext.Load(listCollection);

            //Execute the Query
            clientContext.ExecuteQuery();

            //Get a the global nav list
            List globalNav = listCollection.Where(l => l.Title == LIST_TITLE).FirstOrDefault();

            bool rerExists = false;
            if (globalNav == null)
            {
                //Create the list
                ListCreationInformation newList = new ListCreationInformation();
                newList.Description = "List which holds our Global Navigation";
                newList.TemplateType = (int)ListTemplateType.GenericList;
                newList.Title = LIST_TITLE;

                //Add the list to the web
                clientContext.Web.Lists.Add(newList);

                //Update the web
                clientContext.ExecuteQuery();

                //Create a new field
                FieldCreationInformation field = new FieldCreationInformation(FieldType.Number);
                field.AddToDefaultView = true;
                field.DisplayName = "Nav Order";
                field.InternalName = "NavOrder";
                field.Id = new Guid();
                field.Required = false;

                //Get the newly created list
                globalNav = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
                clientContext.Load(globalNav);
                clientContext.ExecuteQuery();

                globalNav.CreateField(field);

            }
            else
            {
                //Load globalNav 
                clientContext.Load(globalNav);

                //Load EventReceivers
                clientContext.Load(globalNav.EventReceivers);

                //Execute the Query
                clientContext.ExecuteQuery();

                //Execute the Query
                clientContext.ExecuteQuery();

                //Check to see if the event Receiver Exists
                foreach(var receiver in globalNav.EventReceivers)
                {
                    if (receiver.ReceiverName == RECEIVER_NAME)
                    {
                        rerExists = true;
                    }
                }
            }

            if (!rerExists)
            {
                //create the EventReceiverDefinitionCreationInformation object
                EventReceiverDefinitionCreationInformation reciever = new EventReceiverDefinitionCreationInformation();

                //Add the EventType
                reciever.EventType = EventReceiverType.ItemAdded;

                //Get the Receiver Url / WCF Url
                reciever.ReceiverUrl = WCFUrl();

                //Name the Event Receiver
                reciever.ReceiverName = RECEIVER_NAME;

                //synchronization
                reciever.Synchronization = EventReceiverSynchronization.Synchronous;

                //Add Event Receiver to the List
                globalNav.EventReceivers.Add(reciever);

                //Update the ClientContext
                clientContext.ExecuteQuery();

                System.Diagnostics.Trace.WriteLine("Added ItemAdded receiver at " + reciever.ReceiverUrl);
            }
        }

        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            //Get the list
            List globalNav = clientContext.Web.Lists.GetByTitle(LIST_TITLE);

            //Load the Event Receiver
            clientContext.Load(globalNav.EventReceivers);

            //Execute the Query
            clientContext.ExecuteQuery();

            //Load the Event Receiver
            var rer = globalNav.EventReceivers.Where(r => r.ReceiverName == RECEIVER_NAME).FirstOrDefault();

            try
            {
                //Delete Event Receiver
                rer.DeleteObject();

                //Update ClientContext
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
        }

        public void ItemAddedEvent(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                //Get the list
                List globalNav = clientContext.Web.Lists.GetById(listId);

                //Get the Item
                ListItem item = globalNav.GetItemById(listItemId);

                //Execute the Query
                clientContext.ExecuteQuery();

                //Update Column NavOrder
                item["NavOrder"] = "1";

                //Update the item
                item.Update();

                //Execute the Query
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
        }

        private string WCFUrl()
        {
            OperationContext op = OperationContext.Current;
            Message msg = op.RequestContext.RequestMessage;
            return msg.Headers.To.ToString();
        }
    }
}