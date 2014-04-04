using System;
using System.Collections.Generic;
using System.DirectoryServices;

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

        public List<Member> BuildGroupMembers(MailGroup mailGroup)
        {
            throw new NotImplementedException();
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
