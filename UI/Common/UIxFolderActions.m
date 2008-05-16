/* UIxFolderActions.m - this file is part of SOGo
 *
 * Copyright (C) 2007 Inverse groupe conseil
 *
 * Author: Wolfgang Sourdeau <wsourdeau@inverse.ca>
 *
 * This file is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2, or (at your option)
 * any later version.
 *
 * This file is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; see the file COPYING.  If not, write to
 * the Free Software Foundation, Inc., 59 Temple Place - Suite 330,
 * Boston, MA 02111-1307, USA.
 */

#import <Foundation/NSDictionary.h>
#import <Foundation/NSString.h>
#import <Foundation/NSUserDefaults.h>
#import <Foundation/NSURL.h>
#import <Foundation/NSValue.h>

#import <NGObjWeb/WOContext.h>
#import <NGObjWeb/WOContext+SoObjects.h>
#import <NGObjWeb/WORequest.h>
#import <NGObjWeb/WOResponse.h>
#import <NGObjWeb/SoSecurityManager.h>
#import <NGObjWeb/SoClassSecurityInfo.h>

#import <SoObjects/SOGo/LDAPUserManager.h>
#import <SoObjects/SOGo/NSArray+Utilities.h>
#import <SoObjects/SOGo/SOGoGCSFolder.h>
#import <SoObjects/SOGo/SOGoPermissions.h>
#import <SoObjects/SOGo/SOGoUser.h>

#import "WODirectAction+SOGo.h"

#import "UIxFolderActions.h"

@implementation UIxFolderActions

- (void) _setupContext
{
  NSString *mailInvitationParam;
  SOGoUser *activeUser;

  activeUser = [context activeUser];
  login = [activeUser login];
  clientObject = [self clientObject];
  owner = [clientObject ownerInContext: nil];

  baseFolder = [[clientObject container] nameInContainer];

  um = [LDAPUserManager sharedUserManager];
  ud = [activeUser userSettings];
  moduleSettings = [ud objectForKey: baseFolder];
  if (!moduleSettings)
    {
      moduleSettings = [NSMutableDictionary new];
      [moduleSettings autorelease];
    }
  [ud setObject: moduleSettings forKey: baseFolder];

  mailInvitationParam
    = [[context request] formValueForKey: @"mail-invitation"];
  isMailInvitation = [mailInvitationParam boolValue];
}

- (WOResponse *) subscribeAction
{
  [self _setupContext];

  return [clientObject subscribe: YES
		       inTheNameOf: nil
		       fromMailInvitation: isMailInvitation
		       inContext: context];
}

- (WOResponse *) unsubscribeAction
{
  [self _setupContext];

  return [clientObject subscribe: NO
		       inTheNameOf: nil
		       fromMailInvitation: isMailInvitation
		       inContext: context];
}

- (WOResponse *) canAccessContentAction
{
  /* We want this action to be authorized managed by the SOPE's internal acl
     handling. */
  return [self responseWith204];
// #warning IMPROVEMENTS REQUIRED!
//   NSArray *acls;
// //  NSEnumerator *userAcls;
// //  NSString *currentAcl;

//   [self _setupContext];
  
// //  NSLog(@"canAccessContentAction %@, owner %@", subscriptionPointer, owner);

//   if ([login isEqualToString: owner] || [owner isEqualToString: @"nobody"]) {
//     return [self responseWith204];
//   }
//   else {
//     acls = [clientObject aclsForUser: login];
// //    userAcls = [acls objectEnumerator];
// //    currentAcl = [userAcls nextObject];
// //    while (currentAcl) {
// //      NSLog(@"ACL login %@, owner %@, folder %@: %@",
// //	    login, owner, baseFolder, currentAcl);
// //      currentAcl = [userAcls nextObject];
// //    }
//     if (([[clientObject folderType] isEqualToString: @"Contact"]
// 	 && [acls containsObject: SOGoRole_ObjectViewer]) ||
// 	([[clientObject folderType] isEqualToString: @"Appointment"]
// 	 && [acls containsObject: SOGoRole_AuthorizedSubscriber])) {
//       return [self responseWith204];
//     }
//   }
  
//   return [self responseWithStatus: 403];
}

- (WOResponse *) _realFolderActivation: (BOOL) makeActive
{
  NSMutableArray *folderSubscription;
  NSString *folderName;

  [self _setupContext];
  folderSubscription
    = [moduleSettings objectForKey: @"InactiveFolders"];
  if (!folderSubscription)
    {
      folderSubscription = [NSMutableArray array];
      [moduleSettings setObject: folderSubscription forKey: @"InactiveFolders"];
    }

  folderName = [clientObject nameInContainer];
  if (makeActive)
    [folderSubscription removeObject: folderName];
  else
    [folderSubscription addObjectUniquely: folderName];

  [ud synchronize];

  return [self responseWith204];
}

- (WOResponse *) activateFolderAction
{
  return [self _realFolderActivation: YES];
}

- (WOResponse *) deactivateFolderAction
{
  return [self _realFolderActivation: NO];
}

- (WOResponse *) deleteFolderAction
{
  WOResponse *response;

  response = (WOResponse *) [[self clientObject] delete];
  if (!response)
    response = [self responseWith204];

  return response;
}

- (WOResponse *) renameFolderAction
{
  WOResponse *response;
  NSString *folderName;

  folderName = [[context request] formValueForKey: @"name"];
  if ([folderName length] > 0)
    {
      clientObject = [self clientObject];
      [clientObject renameTo: folderName];
      response = [self responseWith204];
    }
  else
    {
      response = [self responseWithStatus: 500];
      [response appendContentString: @"Missing 'name' parameter."];
    }

  return response;
}

- (id) batchDeleteAction
{
  WOResponse *response;
  NSString *idsParam;
  NSArray *ids;

  idsParam = [[context request] formValueForKey: @"ids"];
  ids = [idsParam componentsSeparatedByString: @"/"];
  if ([ids count])
    {
      clientObject = [self clientObject];
      [clientObject deleteEntriesWithIds: ids];
      response = [self responseWith204];
    }
  else
    {
      response = [self responseWithStatus: 500];
      [response appendContentString: @"At least 1 id required."];
    }
  
  return response;
}

@end
