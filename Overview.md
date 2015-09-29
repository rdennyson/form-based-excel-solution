There have been lots of discussion regarding how to operate the excel via windows form, and our solution is based on many open discussions already.

However, we found these solutions cannot support the function to open multiple excel workbooks in the windows form.

Our solution can support this, thus it's the enhancement to these existing solutions and we would like to share with people.

The technical point here is, open excel file as different form tab, and invoke IOleDocumentView::UIActivate to activate each tab when the excel file is clicked to enable the toolbar of current excel file