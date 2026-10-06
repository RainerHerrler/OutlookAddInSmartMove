using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMove.OutlookAddIn
{
    internal static class AddressHeuristic
    {
        private const string RecipientSmtpProperty = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public static bool IsOwnSender(string senderAddress, ISet<string> ownAddresses)
        {
            string normalizedAddress = Normalize(senderAddress);
            return !string.IsNullOrEmpty(normalizedAddress) &&
                ownAddresses != null &&
                ownAddresses.Contains(normalizedAddress);
        }

        public static string Normalize(string address)
        {
            return string.IsNullOrWhiteSpace(address) ? string.Empty : address.Trim().ToLowerInvariant();
        }

        public static IReadOnlyCollection<string> GetToAddresses(Outlook.MailItem mailItem)
        {
            var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Recipients recipients = null;
            try
            {
                recipients = mailItem.Recipients;
                for (int index = 1; index <= recipients.Count; index++)
                {
                    Outlook.Recipient recipient = null;
                    Outlook.PropertyAccessor accessor = null;
                    Outlook.AddressEntry entry = null;
                    Outlook.ExchangeUser exchangeUser = null;
                    try
                    {
                        recipient = recipients[index];
                        if (recipient.Type != (int)Outlook.OlMailRecipientType.olTo)
                        {
                            continue;
                        }

                        string address = string.Empty;
                        try
                        {
                            accessor = recipient.PropertyAccessor;
                            address = accessor.GetProperty(RecipientSmtpProperty) as string;
                        }
                        catch (COMException)
                        {
                            // Some recipient types do not expose the SMTP property.
                        }

                        if (string.IsNullOrWhiteSpace(address))
                        {
                            try
                            {
                                entry = recipient.AddressEntry;
                                if (entry != null)
                                {
                                    exchangeUser = entry.GetExchangeUser();
                                    address = exchangeUser?.PrimarySmtpAddress;
                                }
                            }
                            catch (COMException)
                            {
                                // Fall back to the recipient's regular address.
                            }
                        }

                        if (string.IsNullOrWhiteSpace(address))
                        {
                            address = recipient.Address;
                        }

                        address = Normalize(address);
                        if (address.Contains("@"))
                        {
                            addresses.Add(address);
                        }
                    }
                    catch (COMException)
                    {
                        // An unresolved recipient must not prevent the others from being counted.
                    }
                    finally
                    {
                        ReleaseCom(exchangeUser);
                        ReleaseCom(entry);
                        ReleaseCom(accessor);
                        ReleaseCom(recipient);
                    }
                }
            }
            finally
            {
                ReleaseCom(recipients);
            }

            return addresses;
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
