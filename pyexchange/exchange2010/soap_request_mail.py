"""
(c) 2013 LinkedIn Corp. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License");?you may not use this file except in compliance with the License. You may obtain a copy of the License at  http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software?distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
"""

from .soap_request import M, T, DISTINGUISHED_IDS

import sys
if sys.version_info[0] == 3:
    unicode = str
    basestring = str


def get_mail_items(folder_id='inbox', format=u"Default", query=None, max_entries=1000):
    """
      :param str format: IdOnly or Default or AllProperties
                         - see `doc <http://msdn.microsoft.com/en-us/library/aa580545%28v=exchg.140%29.aspx>`_
      :param str query: advanced query, example: `HasAttachments:true Subject:'Message with Attachments' Kind:email`
                        - see `doc <http://msdn.microsoft.com/en-us/library/ee693615%28v=exchg.140%29.aspx>`_
      :param max_entries: defaults to 1000 as exchange won't return more under default settings
    """
    # http://msdn.microsoft.com/en-us/library/aa566107%28v=exchg.140%29.aspx         FindItem operation
    # http://msdn.microsoft.com/en-us/library/aa566370%28v=exchg.140%29.aspx         FindItem XML
    # http://msdn.microsoft.com/en-us/library/office/dn579420%28v=exchg.150%29.aspx  FindItem + AQS
    xml_fid = T.DistinguishedFolderId(Id=folder_id) if folder_id in DISTINGUISHED_IDS else T.FolderId(Id=folder_id)
    root = M.FindItem(
        {u'Traversal': u'Shallow'},
        M.ItemShape(
            T.BaseShape(format),
            T.AdditionalProperties(
                T.FieldURI(FieldURI=u'item:Subject'),
                T.FieldURI(FieldURI=u'item:DateTimeReceived'),
                T.FieldURI(FieldURI=u'item:Size'),
                T.FieldURI(FieldURI=u'item:Importance'),
                # T.FieldURI(FieldURI=u'item:Attachments'),
                T.FieldURI(FieldURI=u'message:IsRead'),
                )
        ),
        # <FractionalPageItemView MaxEntriesReturned="" Numerator="" Denominator=""/>
        # <m:IndexedPageItemView MaxEntriesReturned="10" Offset="0" BasePoint="Beginning" />
        M.IndexedPageItemView(MaxEntriesReturned=unicode(max_entries), Offset="0", BasePoint="Beginning"),
        M.SortOrder(
            T.FieldOrder(
                T.FieldURI(FieldURI=u'item:DateTimeReceived'),
                Order=u'Descending'
            )
        ),
        M.ParentFolderIds(xml_fid),
        M.QueryString(query) if query else ' '
    )
    return root


def get_attachment(aid):
    """
    The GetAttachment operation is used to retrieve existing attachments on items in the Exchange store.
    See `doc <http://msdn.microsoft.com/en-us/library/aa494316%28v=exchg.140%29.aspx>`
    
    Can raise FailedExchangeException: Exchange Fault (ErrorInvalidIdNotAnItemAttachmentId) from Exchange server
    """
    root = M.GetAttachment(
        M.AttachmentShape(),
        M.AttachmentIds(
            T.AttachmentId(Id=aid)
        )
    )
    return root


