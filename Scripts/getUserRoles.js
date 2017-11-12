function getUserRoles() {
    var roles = Xrm.Page.context.getUserRoles();
    var flag = false;
    for (var i = 0; i < roles.length; i++) {
      if ((roles[i]=="48AE0A84-66C4-E711-A828-000D3AB6BFFF")||(roles[i]=="3AB40A84-66C4-E711-A828-000D3AB6BFFF"))
      {
         flag=true;
      }
    }
    return flag;
}