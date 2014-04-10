using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Runtime.InteropServices;

namespace ReadEmailsFromAD
{
    class ActiveDirectory
    {
        public List<MailGroup> GetAllMailGroups()
        {
            var resultTable = new List<MailGroup>();

            string filter =
                String.Format(
                    "(&(objectClass=Group)(objectCategory=Group)(|(msExchHideFromAddressLists=FALSE)(!(msExchHideFromAddressLists=*))))");

            SearchResultCollection searchResults = FindAll(filter, new[] { "displayName", "mail", "mailNickname", "name" });
            if (searchResults == null)
            {
                return resultTable;
            }

            for (int i = 0; i < searchResults.Count; i++)
            {
                var searchResult = searchResults[i];

                string propDisplayName = GetValue(searchResult, "displayName");
                string propEmailAddress = GetValue(searchResult, "mail");
                string propAliasName = GetValue(searchResult, "mailNickname");
                string propName = GetValue(searchResult, "name");

                if (!String.IsNullOrEmpty(propDisplayName))
                {
                    resultTable.Add(new MailGroup
                    {
                        DisplayName = propDisplayName,
                        EmailAddress = propEmailAddress,
                        AliasName = propAliasName,
                        Name = propName
                    });
                }
            }

            return resultTable;
        }

        public List<Member> BuildGroupMembers(string groupName)
        {
            string filter =
                String.Format(
                    "(&(objectClass=Group)(objectCategory=Group)(|(msExchHideFromAddressLists=FALSE)(!(msExchHideFromAddressLists=*)))(name={0}))",
                    groupName);

            var results = FindOne(filter, new[] { "member" });

            if (results == null)
            {
                return new List<Member>();
            }

            var resultTable = new List<Member>();

            DirectoryEntry directoryEntry = null;
            DirectorySearcher groupSearcher = null;

            foreach (string strProp in results.Properties["member"])
            {
                string path = strProp;
                try
                {
                    if (directoryEntry == null)
                    {
                        directoryEntry = new DirectoryEntry
                        {
                            AuthenticationType = AuthenticationTypes.FastBind
                        };
                    }
                    directoryEntry.Path = @"LDAP://" + path;
                    directoryEntry.RefreshCache();

                    if (groupSearcher == null)
                    {
                        groupSearcher = new DirectorySearcher(directoryEntry)
                        {
                            SearchScope = SearchScope.Base,
                            Filter = "(|(objectClass=group)(objectClass=user))"
                        };
                        groupSearcher.PropertiesToLoad.AddRange(new[] { "displayname", "mail", "mailnickname", "name", "objectClass" });
                    }

                    SearchResult searchResult = groupSearcher.FindOne();

                    if (searchResult == null)
                    {
                        continue;
                    }

                    var objectClassPropertyCollection = searchResult.Properties["objectClass"];
                    if (objectClassPropertyCollection == null)
                    {
                        continue;
                    }

                    if (objectClassPropertyCollection.Contains("user"))
                    {
                        var member = new Member
                        {
                            DisplayName = GetValue(searchResult, "displayname"),
                            EmailAddress = GetValue(searchResult, "mail"),
                            UserID = GetValue(searchResult, "mailnickname")
                        };
                        resultTable.Add(member);
                    }
                    else if (objectClassPropertyCollection.Contains("group"))
                    {
                        var members = BuildGroupMembers(GetValue(searchResult, "name"));
                        foreach (var m in members)
                        {
                            resultTable.Add(m);
                        }
                    }
                }
                catch (COMException)
                {
                }
            }

            if (groupSearcher != null)
            {
                groupSearcher.Dispose();
            }

            if (directoryEntry != null)
            {
                directoryEntry.Close();
            }
            return resultTable;
        }

        private SearchResultCollection FindAll(string filter, string[] propertiesToLoad)
        {
            DirectorySearcher searcher = null;
            try
            {
                searcher = CreateSearcher(filter, propertiesToLoad);
                return searcher.FindAll();
            }
            finally
            {
                if (searcher != null)
                {
                    if (searcher.SearchRoot != null)
                    {
                        searcher.SearchRoot.Dispose();
                    }
                    searcher.Dispose();
                }
            }
        }

        private SearchResult FindOne(string filter, string[] propertiesToLoad)
        {
            DirectorySearcher searcher = null;
            try
            {
                searcher = CreateSearcher(filter, propertiesToLoad);
                return searcher.FindOne();
            }
            finally
            {
                if (searcher != null)
                {
                    if (searcher.SearchRoot != null)
                    {
                        searcher.SearchRoot.Dispose();
                    }
                    searcher.Dispose();
                }
            }
        }

        private DirectorySearcher CreateSearcher(string filter, string[] propertiesToLoad)
        {
            var searcher = new DirectorySearcher
            {
                SearchRoot = new DirectoryEntry(),
                SearchScope = SearchScope.Subtree,
                Filter = filter
            };
            if (propertiesToLoad != null)
            {
                searcher.PropertiesToLoad.AddRange(propertiesToLoad);
            }

            return searcher;
        }

        private static string GetValue(SearchResult searchResult, string propertyName)
        {
            if (searchResult.Properties[propertyName] == null || searchResult.Properties[propertyName].Count == 0)
            {
                return String.Empty;
            }

            object prop = searchResult.Properties[propertyName][0];

            var value = prop as string;

            return value ?? prop.ToString();
        }
    }
}
