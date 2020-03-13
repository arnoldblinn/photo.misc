using System;
using System.Collections.Generic;
using System.Text;

namespace Msn.PhotoMix
{
	public enum PhotoMixError
	{
        // access level errors
        [PhotoMixErrorAttribute(LogError = false, Message = "Invalid partner specified.")]
        InvalidPartner = 1001,
        [PhotoMixErrorAttribute(LogError = false, Message = "Certificate not valid for the specified partner.")]
        InvalidPartnerCert,
        [PhotoMixErrorAttribute(LogError = false, Message = "Partner not authorized for this operation.")]
        PartnerNotAuthorized,
        [PhotoMixErrorAttribute(LogError = false, Message = "Not authorized for this operation.")]
        MemberNotAuthorized,
        [PhotoMixErrorAttribute(LogError = false, Message = "Certificate not valid for this operation.")]
        InvalidManagementCertificate,
        [PhotoMixErrorAttribute(LogError = false, Message = "User has not been authenticated.")]
        MemberNotAuthenticated,

        // domain level errors
        [PhotoMixErrorAttribute(LogError = false, Message = "Invalid domain name.")]
        InvalidDomainName = 2001,
        [PhotoMixErrorAttribute(LogError = false, Message = "Domain name is blocked.")]
        BlockedDomainName,
        [PhotoMixErrorAttribute(LogError = false, Message = "Invalid PhotoMixConfigId was specified.")]
        InvalidPhotoMixConfigId,
        [PhotoMixErrorAttribute(LogError = false, Message = "The domain has not been reserved.")]
        DomainNotReserved,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified domain is not available.")]
        DomainUnavailable,
        [PhotoMixErrorAttribute(LogError = false, Message = "The operation is not valid at this time. The domain is pending changes.")]
        DomainPendingChanges,
        [PhotoMixErrorAttribute(LogError = false, Message = "The operation is not valid at this time. The domain is suspended.")]
        DomainSuspended,
        [PhotoMixErrorAttribute(LogError = false, Message = "The operation is not valid at this time. The domain is pending configuration.")]
        DomainPendingConfiguration,
        [PhotoMixErrorAttribute(LogError = false, Message = "The operation is not permitted on this domain.")]
        NotPermittedForDomain,

        // member level errors
        [PhotoMixErrorAttribute(LogError = false, Message = "Invalid member name.")]
        MemberNameInvalid = 3001,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is not permitted.")]
        MemberNameBlocked,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is not available.")]
        MemberNameUnavailable,
        [PhotoMixErrorAttribute(LogError = false, Message = "Blank membername is not permitted.")]
        MemberNameBlank,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name had invalid characters.")]
        MemberNameIncludesInvalidChars,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name had invalid characters.")]
        MemberNameIncludesDots,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is in use.")]
        MemberNameInUse,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is in use.")]
        ManagedMemberExists,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name does not exist.")]
        ManagedMemberNotExists,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is in use.")]
        UnmanagedMemberExists,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified member name is not an unmanaged user in the domain")]
        UnmanagedMemberNotExists,
        [PhotoMixErrorAttribute(LogError = false, Message = "The maximum number of member accounts has been reached for this domain.")]
        MaxMembershipLimit,
        [PhotoMixErrorAttribute(LogError = false, Message = "A password must be specified.")]
        PasswordBlank,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified password is too short.")]
        PasswordTooShort,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified password is too long.")]
        PasswordTooLong,
        [PhotoMixErrorAttribute(LogError = false, Message = "The password must not include the member name.")]
        PasswordIncludesMemberName,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified password includes invalid characters.")]
        PasswordIncludesInvalidChars,
        [PhotoMixErrorAttribute(LogError = false, Message = "The specified password is invalid.")]
        PasswordInvalid,
        [PhotoMixErrorAttribute(LogError = false, Message = "The net id is invalid.")]
        InvalidNetId,
        [PhotoMixErrorAttribute(LogError = false, Message = "The offer specified is invalid.")]
        InvalidOffer,

        // generic errors
        [PhotoMixErrorAttribute(LogError = true, Message = "Internal error.")]
        InternalError = 9001,
        [PhotoMixErrorAttribute(LogError = false, Message = "Invalid Parameter.")]
        InvalidParameter,
        [PhotoMixErrorAttribute(LogError = true, Message = "Passport error.")]
        PassportError,
        [PhotoMixErrorAttribute(LogError = true, Message = "Exchange error.")]
        ExchangeError,
        [PhotoMixErrorAttribute(LogError = true, Message = "Subscription Services error.")]
        SubscriptionServicesError,
        [PhotoMixErrorAttribute(LogError = true, Message = "Error forced by testing.")]
        TestForcedError,
        [PhotoMixErrorAttribute(LogError = false, Message = "Service down.")]
        ServiceDown,

        // NYI
			[PhotoMixErrorAttribute(LogError = false, Message = "This feature isn't available.")]
		NYI = 10001,

	}
}

