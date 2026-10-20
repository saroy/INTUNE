<#
.SYNOPSIS
    Interactive WinForms wizard that bulk-adds Microsoft Intune assignments and
    scope tags to supported apps, app configurations, or configuration profiles.

.DESCRIPTION
    A tab-based GUI wizard that walks an administrator through:

        Tab 1 (Connect)     - Optionally enter a Tenant ID and Client ID, choose
                              what to work with (apps, app configurations,
                              configuration profiles,
                              or both) and what to change
                              (assignments, scope tags, or both), then
                              authenticate to Microsoft Graph via Connect-MgGraph
                              (delegated). Permissions are checked on connect;
                              Next is enabled once the session is established.
                              Tabs that cannot affect the chosen change are
                              removed from the wizard.
        Tab 2 (Assignment)  - Search Entra ID for the target group, choose
                              Include/Exclude, choose Add/Remove, and optionally
                              apply an assignment filter. Skipped entirely when
                              only scope tags are being changed, in which case
                              Add/Remove is chosen on the Scope Tags tab instead.
        Tab 3 (Scope Tags)  - Optionally select one or more role scope tags to
                              apply to every selected object. The tab follows the
                              action chosen on the Assignment tab: tags are added
                              under Add and removed under Remove. Tags that are
                              not selected are never touched, and an object's
                              last remaining tag is never removed.
        Tab 4 (Objects)     - For Add: search by display-name prefix, narrow the loaded
                              results with free-text, object-type, and platform
                              filters, and multi-select one or more apps,
                              app configurations,
                              configuration profiles, or compliance policies
                              (checkbox grid). Selections are kept across filter
                              changes, so several filters can be used in turn to
                              build up one selection.
                              For Remove: the wizard lists the objects the
                              change applies to, pre-selected. Removing a group
                              means surveying every supported object's
                              assignments; removing scope tags does not, because
                              tag membership comes back with the object list
                              itself. Objects where the group is an exclusion, is
                              the only assignment, or where the selected tags are
                              the only tags, are marked. No prefix search is
                              involved.
        Tab 5 (Summary)     - Review, optionally preview the impact, then
                              execute.

    The wizard never replaces an object's assignment set or its scope tag list.
    Every operation is a targeted modification: the requested change is applied
    and everything else already on the object is carried through unchanged.

    Concretely, adding an assignment leaves every existing assignment in place;
    removing one takes out only the selected group and preserves the rest,
    including All Devices, All Users, and Configuration Manager collection
    targets it cannot itself create; and scope tags are only ever added to the
    tags an object already has. Nothing is ever cleared wholesale, and no
    operation touches a scope tag or an assignment that was not asked for.

    All Graph calls use Invoke-MgGraphRequest, which rides the existing
    Connect-MgGraph session. Only Microsoft.Graph.Authentication is required -
    there is no dependency on the typed Microsoft.Graph.Beta.* sub-modules or
    the assembly version-skew crashes they cause.

    WHAT A RUN CHANGES
        Two independent choices are made on the Connect tab. "What do you want to
        work with" selects the object types in scope: apps, app configurations,
        configuration profiles, including compliance policies, or both. "What do you want to change" selects the
        operation: assignments, scope tags, or both.

        Scope tags only is the case with no group at all. The Assignment tab is
        removed from the wizard, Add/Remove moves to the Scope Tags tab, and no
        object's assignments are read or written. For a removal, the objects to
        act on are found by the tags they already carry - which costs nothing
        extra, because roleScopeTagIds is a plain property returned by the same
        list call that enumerates the objects. Finding objects by GROUP is the
        expensive case, since assignments are a relationship and have to be read
        per object.

    SCOPE - IOS, ANDROID, MACOS, AND WINDOWS
        Apps, app configuration policies, and device configuration policies are
        listed for iOS, Android, macOS, and Windows across the supported Graph
        object types:
            mobileAppConfigurations  - App configuration policies
            configurationPolicies      - Settings catalog and templates
            deviceConfigurations       - Legacy device configuration profiles
            groupPolicyConfigurations  - ADMX / administrative templates
            deviceCompliancePolicies   - Compliance policies
        Assignment filters are limited to platform windows10AndLater.

    ASSIGNMENT MECHANICS - THIS IS THE IMPORTANT PART
        Apps expose a real assignments collection, so an app assignment is a
        single additive POST to .../mobileApps/{id}/assignments. Nothing else on
        the app is touched.

        Configuration profiles do NOT. Profile types only expose a bulk
        /assign action, and that action REPLACES the entire assignment set with
        whatever is posted. To stay additive, every profile goes through a
        read-merge-write cycle:

            GET  .../{kind}('{id}')/assignments      Read the current set
            (merge the new target into the existing targets, de-duplicated)
            POST .../{kind}('{id}')/assign           Write the merged set back

        If the GET fails for any reason, the POST is not attempted for that
        profile. This is what prevents the classic accident of wiping every
        existing group from a profile.

    SCOPE TAG MECHANICS
        roleScopeTagIds is a full array on the object, so it is also a
        read-merge-write: GET the object, union the existing tag IDs with the
        selected ones, then PATCH the union back. Tags are only ever added.
        The built-in Default scope tag is ID "0".

.EXAMPLE
    .\Update-IntuneAssignment_GUI_Wizard_v11.ps1

    Launches the wizard. Choose a mode and authenticate on Tab 1, then proceed
    through the tabs.

.EXAMPLE
    Add a group with an assignment filter to a set of settings catalog policies:

        1. Tab 1  - Mode = Configuration Profiles, Authenticate.
        2. Tab 2  - Type a name prefix, click Load, check the target policies.
        3. Tab 3  - Search for the group, select it, leave Action = Add,
                    pick the filter and set the filter mode to Include.
        4. Tab 4  - Leave empty to make no scope tag changes.
        5. Tab 5  - Click Preview Impact to confirm existing assignment counts,
                    then Finish.

.NOTES
    Author : John Marcum (PJM) @PJ_Marcum
    Version: 11.0

    Change history:
        11.0- Renamed from Add-IntuneAssignment to Update-IntuneAssignment. The
              tool has removed assignments and scope tags since 9.1, so "Add" had
              become misleading in the one place a reader looks first. "Update"
              is an approved PowerShell verb and covers both directions. The log
              file follows the script name and is now
              Update-IntuneAssignment.log; any existing Add-IntuneAssignment.log
              is left in place rather than renamed or merged, so earlier session
              history is preserved where it is.
        10.9- Cleared the pre-filled tenant ID and the default object and group
              search prefixes. All three were specific to one organisation and
              had no business shipping with the script; each remains a variable
              at the top of the file for anyone who wants to set their own.
        10.8- The Tenant and Client ID hint is drawn in the normal text colour,
              which was the last of the grey labels. Grey read as decoration on a
              line that tells the operator both fields are optional and what
              happens if either is left blank.
        10.7- Batched reads now account for themselves in the log. Each attempt
              records how many requests were sent, succeeded, failed outright,
              and were queued for retry, along with how long it took, and the run
              ends with a line reconciling results against requests. Previously a
              retry that succeeded looked identical in the log to one that had
              been quietly dropped - the totals only reconciled by arithmetic
              after the fact. A request that receives no response at all is now
              detected explicitly and marked unreadable rather than being absent
              from the results, since a missing key would read to the caller as
              "no assignments" instead of "not answered".
        10.6- The Connect tab selection hint is drawn in the normal text colour
              as well. It states the consequence of the pairing chosen just above
              it - including whether the run involves a survey - which is the
              kind of thing that should be read before continuing rather than
              styled as an aside.
        10.5- The Remove survey note is drawn in the normal text colour instead
              of the grey used for incidental hints. It explains why the next
              step is slow and what it will cost, so it is meant to be read
              rather than skimmed past as fine print.
        10.4- Every wizard step is now logged. Tab moves go through one helper
              that records the direction and both tabs, and each tab records the
              decision made on it: the group and its assignment settings, the
              scope tags and the direction they will be applied in, and the
              object selection broken down by type. A session can now be
              reconstructed from the log alone, which is the first thing wanted
              when a change looks wrong after the fact.
        10.3- Made a mode change after sign-in observable and safe. Next already
              re-read the selectors, but nothing recorded the result: the only
              log line naming the mode was written at authentication, so a later
              change could not be confirmed from the log. Next now logs the mode
              and change scope it is proceeding with. Because the Intune RBAC
              check had run against the mode selected at sign-in, changing the
              mode now repeats that check against the new resource family - the
              effective permissions are cached, so it costs nothing - and warns
              if the role does not cover it. Changing either selector also clears
              any object list built under the previous ones.
        10.2- Fixed the Connect tab selectors having no effect after the first
              sign-in. Both dropdowns were read only inside the Authenticate
              handler, so changing one afterwards left the run using whatever had
              been selected when the session was established - a Remove in
              profiles mode would survey the app catalogue instead, reporting the
              same object count whichever option was chosen. They are now
              resolved whenever they change and again on Next. Also fixed the
              two-line non-replacement notice overlapping the log path line.
        10.1- Replaced the two radio group boxes on the Connect tab with two
              dropdowns, which reclaims roughly 150 vertical pixels on what had
              become the densest tab. Kept as two lists rather than one combined
              list: the axes are independent, so combining them would have
              produced nine entries to describe three choices. The explanatory
              text that sat under each radio is now a single hint line that
              rewrites itself for the pairing actually selected, so it says what
              this run will do instead of listing every possibility.
        10.0- Added a second selector on the Connect tab - "What do you want to
              change": assignments only, scope tags only, or both - matching the
              existing object-type selector. Scope tags could previously only be
              changed as a side effect of a group operation, so removing a tag
              without touching a group was impossible. Choosing scope tags only
              removes the Assignment tab from the wizard, moves Add/Remove to the
              Scope Tags tab, and skips all assignment reads and writes. Its
              removal flow finds objects by the tags they already carry, which
              needs no survey: roleScopeTagIds is a plain property returned by
              the same list call that enumerates objects, unlike assignments,
              which are a relationship and must be read per object. Choosing
              assignments only removes the Scope Tags tab. Summary,
              confirmation, and preview all state which halves are in play.
        9.9 - Stopped a sorted column's header rendering as though it were
              selected. A DataGridView paints the sorted column's header with the
              grid's selection colours, so clicking a header to sort left it
              highlighted in the same blue as a selected row - misleading on the
              group grid, where exactly one row really is selected and that
              selection matters. Header selection colours now match the normal
              header colours on both grids. Sorting still works.
        9.8 - Going Next and then Back no longer discards the object selection.
              Entering the Objects tab rebuilt the list unconditionally, which in
              Remove mode re-ran the entire survey - minutes on a large tenant -
              and cleared every deselection made on the way through. The list is
              now rebuilt only when the group, action, or mode it was derived
              from has actually changed; returning to the tab with those
              unchanged shows the previous list and selection as they were.
              Changing any of them upstream still invalidates the list, so a
              selection made against a different group can never be carried
              forward silently. Load and Refresh remain deliberate rebuilds.
        9.7 - The Remove note on the Assignment tab now states why the survey is
              needed rather than only that it is: $expand=assignments would
              answer the question in one call, but it is deprecated for apps and
              returns 503 and 504 on the configuration profile endpoints, so each
              object is read individually and the reads are batched. Knowing the
              constraint belongs to the API makes the wait explicable rather than
              looking like a slow tool.
        9.6 - Moved Scope Tags ahead of Objects, so every decision about what to
              change is made before the object list is built. This matters for
              Remove: building that list surveys the estate, which is the one
              long wait in the tool, and asking the remaining questions after it
              stranded the operator mid-task. It also makes the object selection
              the last input, so nothing taken later can invalidate it. Summary
              construction moved into its own function accordingly.
        9.5 - Scope tags now follow the Add/Remove action instead of always being
              added. Previously, selecting scope tags during a removal silently
              ADDED them - one tab quietly doing the opposite of the selected
              action, which was the least defensible of the options. Removal is a
              read-merge-write like every other operation here, so tags that are
              not selected are never touched. An object's last remaining scope
              tag is never removed: that object is skipped and reported, since an
              object with no scope tags is invisible to every scoped
              administrator. The Scope Tags tab, the summary, the confirmation,
              and the result dialog all state which direction is in effect.
        9.4 - Fixed the Remove note on the Assignment tab drawing on top of the
              assignment filter box. Filters play no part in a removal, so the
              filter box is now hidden while Remove is selected and the note
              occupies the space instead of overlapping it. The note and the
              large-survey prompt are also composed for the mode in use: they
              name what will actually be surveyed, and only suggest narrowing the
              mode when "Both" is selected and there is something to narrow.
        9.3 - Made the cost of the Remove action visible. Graph cannot be asked
              which objects target a group, so discovery reads every supported
              object's assignments and the cost scales with the estate rather
              than with the number of objects affected. The Assignment tab now
              says so when Remove is selected, the object list is retrieved
              before the expensive assignment survey so the size is known up
              front, and above a configurable threshold the wizard reports how
              many objects and roughly how many batched requests are involved and
              asks before starting. Survey duration is logged. "Both" mode is
              identified as the slowest path, with the suggestion to narrow the
              mode where that will do.
        9.2 - Reworded the guarantee shown on the Connect tab and stated in the
              description. "Only adds assignments and scope tags" described the
              tool before Remove existed and was no longer true. The property
              that actually holds is that no operation ever replaces an object's
              assignment set or scope tag list - each one applies the requested
              change and carries everything else through unchanged. The
              confirmation dialogs use the same wording.
        9.1 - Reordered the wizard so Assignment precedes Objects, and added
              group-driven removal. The group and the action are the subject of
              the operation, and for Remove the object list cannot sensibly be
              typed in - so choosing Remove now makes the Objects tab survey
              every supported object and list those already targeting the group,
              pre-selected, with no prefix search. The survey uses batched
              requests, and an object whose assignments cannot be read is
              reported rather than treated as not matching. Objects where the
              group is an exclusion, or is the object's only assignment, are
              highlighted in the grid and counted on the summary and the
              confirmation - removing an exclusion widens scope, and removing a
              sole assignment leaves a policy applying to nobody. A third mode,
              "Both", covers apps and configuration profiles together, which is
              what clearing a group from an estate requires; the RBAC check runs
              against each resource family in turn.
        9.0 - Preview Impact now reads assignments through the Graph JSON batch
              endpoint rather than one request per object, cutting a preview of
              many objects from one round trip each to one per twenty. Batching
              is used in preference to $expand=assignments, which is deprecated
              on mobileApps and returns 503 and 504 from the configuration
              profile endpoints - an intermittently short result being the worst
              possible input to anything that later changes assignments. Batched
              requests that are throttled or hit a transient fault are retried
              with backoff individually rather than failing the whole batch, and
              an object whose assignments cannot be read is now reported as
              UNKNOWN instead of appearing to have none.
        8.11- An apparent "user cancelled" sign-in now explains that it is often
              not one. When Entra rejects a request in the browser - a missing
              reply URL, public client flows disabled - MSAL reports only a
              cancellation and the AADSTS code never leaves the browser window,
              so the code mapping added in 8.10 cannot see it. The cancellation
              path now lists the registration faults that produce it. Also fixed
              the Connect tab hint text running behind the permissions checkbox.
        8.10- Sign-in failures now name the registration setting that needs
              changing instead of showing only "Authentication failed". AADSTS
              codes are specific and stable, but the accompanying text describes
              the protocol rather than the fix, and nearly every failure here
              comes from an app registration built for a different flow being
              pointed at an interactive tool. Covers missing reply URL, public
              client flows disabled, redirect mismatch, unknown client or tenant,
              missing consent, and Conditional Access blocks.
        8.9 - Stopped the tool silently escalating an app registration's
              permissions. Requesting an explicit scope list triggers incremental
              consent, so running the wizard against any app the operator happened
              to name would prompt to add Intune read/write to it - and accepting
              grants it permanently. That is the wrong default for a tool whose
              Client ID is a free-text box. A "Do not request new permissions"
              checkbox now requests the /.default scope instead, asking only for
              what the app already holds; it is selected automatically whenever a
              Client ID is supplied. An under-permissioned app now surfaces as a
              missing-scope report rather than a consent prompt, and the report
              says whether the shortfall can be resolved without granting the app
              anything new.
        8.8 - The scopes granted on the token are now logged on every sign-in,
              not only when one is missing. Delegated permissions support dynamic
              consent, so the token's scopes and the app registration's listed
              permissions can differ; recording the granted set makes it possible
              to tell afterwards which was actually in force.
        8.7 - The Intune RBAC check now understands wildcard grants. An account
              with a full Intune Administrator or Global Administrator role comes
              back from getEffectivePermissions with a single "*" entry that
              grants every action, and matching literal resource names alone
              missed it - so the check reported inconclusive against the most
              privileged accounts there are. "*" is now recognised outright, and
              family-level wildcards such as
              "Microsoft.Intune_DeviceConfigurations_*" are matched by testing
              the required action against each granted entry as a pattern, not
              only the reverse.
        8.6 - The Intune RBAC check no longer warns when it cannot interpret the
              result. Intune resource action names vary by resource family and
              have changed over time, so a pattern matching nothing is far more
              likely to mean this tool guessed the name wrong than that the
              operator lacks the permission - and warning an administrator about
              a permission they hold is worse than staying quiet. The check now
              returns an inconclusive result unless the account clearly holds
              actions for the relevant resource family, and it logs the actions
              actually returned so the patterns can be corrected. Authenticating
              no longer jumps straight to the Objects tab: Tab 1 has its own Next
              button, enabled once a session is established, so permission
              warnings can be read first.
        8.5 - Added an Intune RBAC check at sign-in alongside the existing scope
              check. Delegated scopes and Intune role permissions are separate
              gates - a token can carry every scope this tool requests and still
              be refused because the account holds a read-only Intune role, or
              because its role assignment is scoped to tags that exclude the
              objects being changed. deviceManagement/getEffectivePermissions is
              queried for the account's allowed resource actions and checked for
              the assign and update actions the selected mode needs. The check is
              advisory: if the endpoint is unavailable or returns nothing
              recognisable it is skipped rather than raising a false alarm. Both
              checks now report through a single dialog.
        8.4 - Added a permissions check at sign-in. A successful Connect-MgGraph
              does not mean the token carries the scopes this tool needs, and the
              mismatch previously surfaced only as a 403 on the first write,
              after the operator had already worked through every tab. The
              granted scopes are now compared against the required list and any
              shortfall is logged, shown in red on the Connect tab, and raised in
              a dialog that names the missing scopes - and that calls out the
              built-in app registration explicitly when no Client ID was given,
              since many tenants let that app read Intune but not write to it.
              403, 401, and 429 failures during execution now log an actionable
              explanation alongside the raw Graph error, and a 403 is summarised
              in the result dialog.
        8.3 - Client ID and Tenant ID are now optional. Both are optional on
              Connect-MgGraph's user parameter set, and leaving Client ID empty
              falls back to the SDK's built-in Microsoft Graph PowerShell app
              registration, so the wizard runs without a dedicated app
              registration in any tenant that permits that app. Empty values are
              omitted from the call rather than passed as empty strings, which is
              what produced the earlier "Cannot bind argument to parameter
              ClientId" failure. The log now records which app registration and
              which tenant scope were used.
        8.2 - Fixed the intermittent 400 "Property target in payload has a value
              that does not match schema" on /assign. Introduced in 7.2:
              Get-ProfileAssignment returns its array comma-protected to stop
              PowerShell unrolling it, but the call sites ALSO wrapped the result
              in @(). Collecting an output stream whose single item is itself an
              array produces a nested array, so the merge loop ran once with the
              whole array bound to $assignment, and $assignment.target
              member-enumerated into an array of targets - posting several
              targets inside one assignment entry. A profile with exactly one
              existing assignment survived this by coincidence, which is why the
              failure looked random. The @() wrappers are removed, and
              Set-ProfileAssignment now flattens and warns if it is ever handed a
              nested collection again. Also added validation for empty Tenant ID
              and Client ID, which previously failed as a parameter binding
              exception in the log.
        8.1 - Fixed a 400 "Property target in payload has a value that does not
              match schema" when merging assignments onto some profiles. An
              assignment target read from GET .../assignments is not necessarily
              valid to POST back to /assign; the read model can carry properties
              the write model rejects, and which ones varies by target type and
              by how the assignment was originally created. Existing targets are
              now rebuilt from an allow list of the properties valid for their
              own @odata.type instead of being echoed back verbatim, a missing
              @odata.type is inferred, and unrecognised target types are
              preserved unchanged so nothing is silently dropped from the merge.
              A rejected /assign payload is now written to the log, since Graph's
              error names neither the property nor the assignment at fault.
        8.0 - Fixed the Connect tab reporting "Connected as ." with no identity
              after a successful sign-in. Root cause, present since 6.0: the
              helper was named Connect-Graph, which Microsoft.Graph.Authentication
              already exports as an ALIAS for Connect-MgGraph. PowerShell resolves
              aliases ahead of functions, so once the module was loaded every call
              bypassed the script's function entirely and invoked Connect-MgGraph
              with only the ClientId and TenantId whose parameter names happened to
              match. The -Scopes list was never requested (sessions worked only on
              previously consented scopes), and none of the function's logging ran,
              which is why no "Connected to Graph as" line ever reached the log.
              The helper is renamed Connect-GraphSession, and a startup check now
              logs an error if any script function is shadowed by an alias or
              cmdlet. Connect-MgGraph output is also discarded so it cannot
              pollute the return value, -NoWelcome is used where supported, and
              the status line falls back through Account, AppName, and ClientId
              rather than rendering an empty identity, and now shows the tenant.
        7.3 - Added a client-side filter row to the Objects tab: a free-text
              display-name box and an object-type dropdown that narrow the
              already-loaded results without re-querying Graph. Checkbox
              selections are now tracked by object Id and survive filter
              changes, so a selection can be built up across several filters.
              Select All applies to the rows currently visible under the filter.
              The count readout now shows how many objects are visible, how many
              were loaded, and how many are selected.
        7.2 - Fixed the result dialog and the "Done" log line always reporting
              zero. The execution loop ran inside Invoke-WithWaitCursor, which
              invokes its scriptblock in a child scope, so every counter
              increment was discarded. The loop now sets the cursor inline.
              Fixed misreported assignment counts for profiles holding exactly
              one existing assignment: the returned array was unrolled to a bare
              Hashtable, so .Count returned its key count (4) instead of 1. The
              assignments themselves were always merged correctly; only the
              before-count in the log and in Preview Impact was wrong.
        7.1 - Fixed a Graph query failure when loading configuration profiles.
              The deviceCompliancePolicies endpoint rejects startswith() in
              $filter, and the resulting 400 aborted the entire profile load.
              All object queries now attempt the server-side prefix filter and
              transparently fall back to retrieving the collection and matching
              the prefix client-side. Fixed the Connect tab mode group box,
              where the descriptive hint label overlapped and clipped the
              configuration profiles option.
        7.0 - Added Windows configuration profile support (settings catalog,
              legacy device configurations, ADMX, and compliance policies) with
              additive read-merge-write assignment handling. Added assignment
              filter support to both apps and profiles. Added additive role
              scope tag assignment. Group selection is now search-based instead
              of bulk-loading every group in the tenant. Merged the old Prefix
              tab into the Objects tab. Renamed Log-Message to Write-LogMessage
              (approved verb) and renamed the $matches local to avoid shadowing
              the automatic variable. Log path now prefers C:\Windows\Logs and
              falls back to the user TEMP folder when not writable.
        6.4 - Apps list filtered to Win32 app types client-side.
        6.3 - Assignment settings tailored to the app type.
        6.2 - Add intent changed to Available; add is now idempotent; added
              Select All toggle.
        6.1 - Module loading no longer imports Microsoft.Graph.Authentication
              when the Graph cmdlets are already loaded in the session.
        6.0 - Rebuilt as a consistent 5-tab wizard; all Graph calls moved to
              Invoke-MgGraphRequest; added Add/Remove action support.

    Authentication and consent:
        By default this tool requests an explicit list of delegated scopes. That
        triggers INCREMENTAL CONSENT - Entra prompts to grant any requested scope
        the app does not already hold, and accepting permanently adds it to that
        app registration. Pointing the tool at an app created for another purpose
        would therefore escalate it to Intune read/write on first run.

        The "Do not request new permissions" checkbox requests the /.default
        scope instead, which asks only for what the app has already been granted.
        Nothing is ever added. An under-permissioned app simply returns a token
        without the required scopes, and the post-sign-in check reports precisely
        which are missing. The checkbox is selected automatically whenever a
        Client ID is present, and is unavailable without one - the built-in
        Microsoft Graph PowerShell app holds no standing grant of its own and
        must request its scopes explicitly.

    Tenant and client:
        Both Tenant ID and Client ID are optional. With Client ID empty the
        Microsoft Graph PowerShell SDK signs in using its own multi-tenant app
        registration, "Microsoft Graph Command Line Tools", which works in most
        tenants and raises a consent prompt on first use. Supply your own app
        registration instead when that app is blocked or unconsented, when app
        consent policies stop users consenting to it, when Conditional Access
        targets specific applications, or when the Intune permissions need to be
        held by an app the organisation controls and audits. A custom app
        registration must be granted the delegated scopes below, and be
        configured as a public client with a redirect URI of http://localhost.

        With Tenant ID empty the sign-in resolves against the account's home
        tenant; supply it when the account is a guest in more than one tenant.

    Required delegated Graph scopes:
        DeviceManagementApps.ReadWrite.All            Apps and app assignments
        DeviceManagementConfiguration.ReadWrite.All   Profiles, assignments,
                                                      and roleScopeTagIds
        DeviceManagementRBAC.Read.All                 Read the scope tag list
        Group.Read.All                                Search target groups

    Scale of the Remove action:
        Graph exposes no way to ask which objects target a given group, so the
        only way to find them is to read every supported object's assignments.
        Listing a collection with $expand=assignments would return them all in
        one call, but that is deprecated for mobileApps - the documented guidance
        is to list without the expand and then expand individual items - and on
        the configuration profile endpoints it is unreliable, returning 503 and
        504 under load. An intermittently truncated result is the worst possible
        input to an operation that then removes assignments, so the wizard reads
        each object individually and batches those reads instead.
        The cost therefore grows with the size of the estate rather than with the
        number of objects actually affected. Requests are batched twenty at a
        time, which makes a few hundred objects quick and a few thousand a wait
        of some minutes; the window is unresponsive while it runs. "Both" mode is
        the slowest, since it surveys apps and all four profile types - narrowing
        the mode on the Connect tab to just apps or just profiles roughly halves
        the work. Above $global:DiscoveryPromptThreshold objects the wizard
        reports the size of the survey and asks before starting. The survey is
        read-only, so cancelling it changes nothing.

    Limitations / assumptions:
                - Apps and profiles are supported for iOS, Android, macOS, and Windows.
                    Assignment filters are currently limited to platform windows10AndLater.
                    Object types that cannot be mapped to one of the four platforms are
                    logged and skipped.
        - App assignment intent is "available". Available installs are
          user-driven via Company Portal and resolve against the signed-in user,
          so target user groups for predictable results.
        - A group can only be targeted once per object. If the selected group is
          already assigned, the object is skipped rather than duplicated
          (AlreadyExists), and if it is assigned with the opposite include /
          exclude type the object is skipped as a Conflict. Existing filters on
          an already-assigned group are not modified.
        - Remove matches on groupId only and will remove BOTH include and
          exclude assignments for that group. For profiles, Remove rewrites the
          assignment set with the remaining targets preserved.
        - Assignment targets other than groups (All Devices, All Users,
          Configuration Manager collections) that already exist on a profile are
          preserved verbatim during the merge, but cannot be created here.
        - Scope tag changes follow the Add/Remove action. Tags not selected are
          never touched. Removal will not take an object's last remaining scope
          tag: that object is skipped and reported, because an object with no
          scope tags is not visible to any scoped administrator and restoring it
          needs an unscoped account.

    LEGAL DISCLAIMER:
    This script is provided "AS IS" with no warranties, express or implied,
    and confers no rights. Use of any portion is at your own risk. Test
    thoroughly in a non-production environment before use. The author is not
    responsible for any damages arising from its use.
#>

# ---------------------------------------------------------------------------
# Configuration / default values
# ---------------------------------------------------------------------------
# Client ID is OPTIONAL. Leave it empty to use the Microsoft Graph PowerShell
# SDK's own multi-tenant app registration ("Microsoft Graph Command Line Tools"),
# which works in most tenants and prompts for consent on first use. Set it to
# your own app registration when the built-in app is blocked or unconsented, when
# app consent policies prevent users consenting to it, when Conditional Access
# targets specific applications, or when the Intune permissions need to sit on an
# app your organisation controls and audits.
# NOTE: many tenants allow the built-in app to READ Intune but not to WRITE to
# it. That combination looks like a clean sign-in followed by 403 Forbidden on
# every assignment and scope tag change. Set this to your own app registration if
# that is your tenant - the wizard now checks the granted scopes at sign-in and
# warns before you get that far.
$defaultClientId    = ""                                         # Optional app registration (client) ID

# Tenant ID is also optional - omitting it signs in against the account's home
# tenant. Set it when the signing-in account is a guest in more than one tenant.
$defaultTenantId    = ""                                         # Optional tenant ID

# Pre-filled search text for the Objects and Assignment tabs. Both are left empty
# so the script carries no organisation's naming conventions with it; set them to
# the prefixes used in your tenant to save typing on every run.
$defaultObjectPrefix = ""
$defaultGroupSearch  = ""

$global:ClientId       = $defaultClientId
$global:TenantId       = $defaultTenantId
$global:Mode           = "Apps"      # Apps | Profiles
$global:Objects        = @()
$global:SelectedObjects = @()

# Checkbox state is tracked by object Id in a set rather than read off the grid.
# The grid only ever shows the rows matching the current list filter, so reading
# selections from visible rows would silently drop anything the user ticked and
# then filtered out of view.
$global:CheckedObjectIds = New-Object 'System.Collections.Generic.HashSet[string]'

# Identifies the Assignment-tab inputs the current object list was built from.
# Re-entering the Objects tab with the same inputs keeps the list and the
# operator's selection; a change to any of them invalidates both.
$global:DiscoveryContext = ""

# Assignments | ScopeTags | Both - what the run is allowed to change. Kept
# separate from $global:Mode, which decides what kind of object is in scope.
$global:ChangeScope = "Assignments"

# Set while the grid is being re-rendered so programmatic cell writes do not
# feed back into the selection set through CellValueChanged.
$global:SuppressGridEvents = $false
$global:GroupId        = $null
$global:SelectedGroup  = $null
$global:AssignmentType = "Include"   # Include | Exclude (Add action only)
$global:Action         = "Add"       # Add | Remove
$global:FilterId       = $null
$global:FilterName     = "(None)"
$global:FilterMode     = "none"      # none | include | exclude
$global:AssignmentFilters = @()
$global:ScopeTags         = @()
$global:SelectedScopeTags = @()
$global:Groups            = @()

# Log to C:\Windows\Logs when writable (elevated), otherwise the user TEMP
# folder so the tool still works in a normal admin desktop session.
$global:LogPath = "$env:windir\Logs\Update-IntuneAssignment.log"
try {
    $probe = "$env:windir\Logs\Update-IntuneAssignment.probe"
    New-Item -Path $probe -ItemType File -Force -ErrorAction Stop | Out-Null
    Remove-Item -Path $probe -Force -ErrorAction SilentlyContinue
} catch {
    $global:LogPath = "$env:TEMP\Update-IntuneAssignment.log"
}

# Required Graph scopes and the single module dependency
# Above this many objects, Remove discovery asks before starting. Surveying is
# read-only, but it is the longest unattended wait in the tool.
$global:DiscoveryPromptThreshold = 250

$global:RequiredScopes = @(
    "DeviceManagementApps.ReadWrite.All",
    "DeviceManagementConfiguration.ReadWrite.All",
    "DeviceManagementRBAC.Read.All",
    "Group.Read.All"
)

$graphBeta = "https://graph.microsoft.com/beta"
$graphV1   = "https://graph.microsoft.com/v1.0"

# Win32 app types require a dedicated assignment settings payload.
$global:Win32AppTypes = @(
    "#microsoft.graph.win32LobApp",
    "#microsoft.graph.win32CatalogApp"
)

# Legacy deviceConfigurations / deviceCompliancePolicies @odata.type values that
# are Windows but do not begin with "windows".
$global:WindowsExtraOdataTypes = @(
    "sharedPCConfiguration",
    "editionUpgradeConfiguration",
    "secureAssessmentConfiguration"
)

# Load WinForms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Helper functions - logging, auth, and Graph plumbing
# ---------------------------------------------------------------------------

# Write a timestamped message to the console and to the log file.
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "$timestamp [$Level] $Message"
    Write-Host $logEntry
    try {
        Add-Content -Path $global:LogPath -Value $logEntry -ErrorAction SilentlyContinue
    } catch {
        # Logging to file is best-effort; never block the UI on a log failure.
    }
}

# Ensure the Graph authentication cmdlets are available. Returns $true on success.
#
# Only Connect-MgGraph, Get-MgContext, and Invoke-MgGraphRequest are used by this
# script, and those exist across all Microsoft.Graph.Authentication 2.x builds. If
# any version of the module is already loaded into the session, we use it as-is and
# do NOT call Import-Module - importing a different on-disk version over an
# already-loaded assembly throws "Assembly with same name is already loaded".
function Assert-GraphAuthModule {
    $moduleName = "Microsoft.Graph.Authentication"

    # 1. Already usable? Use whatever is loaded and skip the import entirely.
    if (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue) {
        $loaded = Get-Module -Name $moduleName | Select-Object -First 1
        if ($loaded) {
            Write-LogMessage "$moduleName already loaded (version $($loaded.Version)); using existing session."
        } else {
            Write-LogMessage "Graph authentication cmdlets already available; using existing session."
        }
        return $true
    }

    # 2. Not loaded yet - install if missing, then import.
    try {
        $installed = Get-Module -Name $moduleName -ListAvailable |
                     Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            Write-LogMessage "Installing module $moduleName ..."
            Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
            Write-LogMessage "Module installed."
        } else {
            Write-LogMessage "Module $moduleName found (version $($installed.Version))."
        }
        Import-Module $moduleName -ErrorAction Stop
        return $true
    } catch {
        Write-LogMessage "Failed to load $moduleName : $_" "ERROR"
        Write-LogMessage "A conflicting Graph version may already be loaded. Start a new PowerShell session and retry." "ERROR"
        return $false
    }
}

# Build a human-readable identity string for logging and the Connect tab.
#
# Get-MgContext does not always populate Account - it can be empty when the token
# is served from the MSAL cache rather than a fresh interactive sign-in, and it is
# always empty for an app-only session. Falling back through the other identifying
# properties avoids rendering a bare "Connected as ." with nothing in it.
function Get-GraphIdentityLabel {
    param ($Context)

    if (-not $Context) { return "unknown identity" }
    foreach ($candidate in @($Context.Account, $Context.AppName, $Context.ClientId)) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { return "$candidate" }
    }
    return "unknown identity"
}

# Authenticate to Microsoft Graph. Returns the Graph context on success, else $null.
#
# DO NOT rename this back to Connect-Graph. Microsoft.Graph.Authentication exports
# Connect-Graph as an ALIAS for Connect-MgGraph, and PowerShell resolves aliases
# ahead of functions. A function by that name is silently shadowed the moment the
# module is loaded, so every call went straight to Connect-MgGraph with only the
# ClientId and TenantId that happened to match its parameter names - the -Scopes
# in this function were never requested, and none of its logging ever ran.
function Connect-GraphSession {
    param (
        [string]$ClientId,
        [string]$TenantId,
        [switch]$UseGrantedPermissionsOnly
    )
    try {
        # Out-Null is load-bearing. Connect-MgGraph emits a welcome banner and,
        # depending on module version, can write objects to the success stream.
        # Without this they are captured as part of this function's return value,
        # so the caller receives an array whose first element is not the context.
        # A truthy array then passes the caller's "did we connect" check while
        # $context.Account resolves to nothing - which is exactly how the Connect
        # tab ended up showing "Connected as ." on a session that had in fact
        # authenticated correctly. -NoWelcome suppresses the banner where the
        # installed module supports it.
        $connectParams = @{
            ErrorAction = "Stop"
        }

        # Requesting an explicit scope list triggers INCREMENTAL CONSENT: Entra
        # asks the user to grant any requested scope the app does not already
        # hold, and accepting permanently adds it to that app registration. Point
        # this tool at an app created for something else - a reporting app, say -
        # and a single run silently escalates it to Intune read/write.
        #
        # Requesting the /.default scope instead asks only for what the app has
        # already been granted. No new permission is ever added. If the app is
        # under-permissioned the token simply comes back short, and the scope
        # check after sign-in reports exactly which scopes are missing.
        if ($UseGrantedPermissionsOnly) {
            $connectParams.Scopes = @("https://graph.microsoft.com/.default")
            Write-LogMessage "Requesting only permissions already granted to the app (/.default); no consent will be requested for additional scopes."
        } else {
            $connectParams.Scopes = $global:RequiredScopes
            Write-LogMessage "Requesting explicit scopes. Entra will prompt for consent for any this app does not already hold, which ADDS them to the app registration."
        }

        # Both of these are optional on Connect-MgGraph's user parameter set.
        # Passing an empty string is NOT the same as omitting the parameter - it
        # fails argument binding - so only add them when they have a value.
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
            $connectParams.ClientId = $ClientId
            Write-LogMessage "Using app registration $ClientId."
        } else {
            Write-LogMessage "No client ID supplied; using the built-in Microsoft Graph PowerShell app registration."
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $connectParams.TenantId = $TenantId
        } else {
            Write-LogMessage "No tenant ID supplied; signing in against the account's home tenant."
        }

        if ((Get-Command Connect-MgGraph).Parameters.ContainsKey("NoWelcome")) {
            $connectParams.NoWelcome = $true
        }
        Connect-MgGraph @connectParams | Out-Null

        # Select-Object -First 1 guards against the same unrolling problem on the
        # way back out of Get-MgContext.
        $context = Get-MgContext | Select-Object -First 1
        if (-not $context) {
            Write-LogMessage "Connect-MgGraph returned no context." "ERROR"
            return $null
        }

        Write-LogMessage "Connected to Graph as $(Get-GraphIdentityLabel $context) (tenant $($context.TenantId), auth $($context.AuthType))."
        return $context
    } catch {
        Write-LogMessage "Authentication error: $_" "ERROR"
        $global:LastAuthError = "$_"
        $hint = Get-AuthErrorHint -ErrorRecord $_
        if ($hint) {
            Write-LogMessage $hint "ERROR"
            $global:LastAuthHint = $hint
        } else {
            $global:LastAuthHint = $null
        }
        return $null
    }
}

# Retrieve the Intune resource actions the signed-in account is actually allowed
# to perform, via deviceManagement/getEffectivePermissions.
#
# This is a different question from "does the token carry the right scopes".
# Delegated scopes come from the app registration's consent; Intune RBAC comes
# from the role assigned to the account. A token can hold every scope this tool
# requests and still be refused, because the account holds a read-only Intune
# role or its role assignment is scoped to tags that exclude these objects.
#
# Returns a flat list of allowed action strings, or $null when the check could
# not be performed. Never throws - an unavailable endpoint must not block the
# wizard, only leave the operator without the extra warning.
function Get-IntuneEffectivePermission {
    # The function has been exposed under both forms; try each before giving up.
    $candidateUris = @(
        "$graphBeta/deviceManagement/getEffectivePermissions",
        "$graphBeta/deviceManagement/getEffectivePermissions(scope='')"
    )

    foreach ($uri in $candidateUris) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

            $actions = New-Object System.Collections.Generic.List[string]
            foreach ($rolePermission in @($response.value)) {
                # Newer shape: rolePermission.resourceActions[].allowedResourceActions
                foreach ($resourceAction in @($rolePermission.resourceActions)) {
                    foreach ($allowed in @($resourceAction.allowedResourceActions)) {
                        if ($allowed) { $actions.Add("$allowed") }
                    }
                }
                # Older shape: rolePermission.actions[]
                foreach ($legacyAction in @($rolePermission.actions)) {
                    if ($legacyAction) { $actions.Add("$legacyAction") }
                }
            }

            if ($actions.Count -gt 0) {
                Write-LogMessage "Retrieved $($actions.Count) effective Intune resource action(s) for this account."
                return $actions
            }
        } catch {
            Write-LogMessage "Could not read effective Intune permissions from $uri : $_" "WARN"
        }
    }

    Write-LogMessage "Effective Intune permissions could not be determined; skipping the RBAC check." "WARN"
    return $null
}

# Report which of the resource actions this run needs are absent from the
# account's effective Intune permissions.
#
# Returns an object with two properties, deliberately NOT a bare array:
#   Conclusive - $false when the returned vocabulary says nothing about this
#                resource family, so no judgement can be made
#   Missing    - the actions that appear to be absent; empty when all are present
#
# A bare array cannot express this. PowerShell collapses an empty array returned
# from a function into $null, which makes "all required actions present"
# indistinguishable from "could not tell" - and those two must drive opposite
# behaviour.
#
# The inconclusive result matters more than it looks. Intune's resource action
# names have changed over time and differ between resource families, so a
# pattern that matches nothing is far more likely to mean "this tool guessed the
# name wrong" than "the operator lacks the permission". Warning an administrator
# that they cannot do something they can do is worse than staying quiet, so the
# check only speaks up when the account clearly HAS actions for this family and
# the specific assign or update action is missing from them.
function Get-MissingIntunePermission {
    param (
        [string[]]$AllowedActions,
        [string]$Mode
    )

    # "Both" is checked against each family in turn by the caller, so this
    # function only ever sees a single family.
    if ($Mode -eq "Apps") {
        $resourceFamily   = "*MobileApps*"
        $requiredPatterns = [ordered]@{
            "Assign apps" = "*MobileApps*_Assign*"
            "Update apps" = "*MobileApps*_Update*"
        }
        $concreteActions = @{
            "Assign apps" = "Microsoft.Intune_MobileApps_Assign"
            "Update apps" = "Microsoft.Intune_MobileApps_Update"
        }
    } elseif ($Mode -eq "AppConfigurations") {
        $resourceFamily   = "*MobileAppConfigurations*"
        $requiredPatterns = [ordered]@{
            "Assign app configurations" = "*MobileAppConfigurations*_Assign*"
            "Update app configurations" = "*MobileAppConfigurations*_Update*"
        }
        $concreteActions = @{
            "Assign app configurations" = "Microsoft.Intune_MobileAppConfigurations_Assign"
            "Update app configurations" = "Microsoft.Intune_MobileAppConfigurations_Update"
        }
    } else {
        $resourceFamily   = "*DeviceConfigurations*"
        $requiredPatterns = [ordered]@{
            "Assign profiles" = "*DeviceConfigurations*_Assign*"
            "Update profiles" = "*DeviceConfigurations*_Update*"
        }
        $concreteActions = @{
            "Assign profiles" = "Microsoft.Intune_DeviceConfigurations_Assign"
            "Update profiles" = "Microsoft.Intune_DeviceConfigurations_Update"
        }
    }

    # An entry returned by getEffectivePermissions can itself be a wildcard.
    # A full Intune Administrator or Global Administrator comes back with a
    # single "*", which grants every action; family-level wildcards such as
    # "Microsoft.Intune_DeviceConfigurations_*" also occur. Matching literal
    # resource names alone misses both, and reports inconclusive against the
    # most privileged accounts there are.
    if ($AllowedActions -contains "*") {
        Write-LogMessage "Account holds the unrestricted Intune permission '*'; all required resource actions are granted."
        return [pscustomobject]@{ Conclusive = $true; Missing = @() }
    }

    $familyActions = @($AllowedActions | Where-Object { $_ -like $resourceFamily })

    # Log what was actually returned for this family. Without this there is no
    # way to tell a real permissions gap from a pattern that needs updating.
    if ($familyActions.Count -gt 0) {
        Write-LogMessage "Effective Intune actions matching '$resourceFamily': $(($familyActions | Sort-Object -Unique) -join ', ')"
    } else {
        Write-LogMessage "None of the $($AllowedActions.Count) effective Intune action(s) match '$resourceFamily'. Sample of what was returned: $((@($AllowedActions | Sort-Object -Unique) | Select-Object -First 15) -join ', ')"
        return [pscustomobject]@{ Conclusive = $false; Missing = @() }
    }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($label in $requiredPatterns.Keys) {
        $pattern = $requiredPatterns[$label]

        # Two directions, because either side may carry the wildcard:
        #   $_ -like $pattern  - a literal granted action matching what is needed
        #   $concrete -like $_ - a granted WILDCARD entry covering what is needed
        $concrete = $concreteActions[$label]
        $match    = @($familyActions | Where-Object { ($_ -like $pattern) -or ($concrete -like $_) })
        if ($match.Count -eq 0) { $missing.Add("$label ($pattern)") }
    }

    return [pscustomobject]@{ Conclusive = $true; Missing = $missing.ToArray() }
}

# Translate an Entra sign-in failure into the registration change that fixes it.
#
# AADSTS codes are stable and specific, but the accompanying text describes the
# protocol rather than the setting to change. Almost every failure here comes
# from an app registration that was built for a different flow - a confidential
# client with a secret, say - being pointed at an interactive tool.
function Get-AuthErrorHint {
    param ($ErrorRecord)

    $text = "$ErrorRecord"

    if ($text -match "AADSTS500113") {
        return "AADSTS500113 - the app registration has no reply (redirect) URL. Interactive sign-in needs one. In Entra, open the app registration, then Authentication > Add a platform > Mobile and desktop applications, and add http://localhost. Also set 'Allow public client flows' to Yes."
    }
    if ($text -match "AADSTS7000218" -or $text -match "public client") {
        return "AADSTS7000218 - the app registration is not enabled for public client flows. In Entra, open the app registration, then Authentication > Advanced settings, and set 'Allow public client flows' to Yes. A desktop tool cannot hold a client secret, so it must authenticate as a public client."
    }
    if ($text -match "AADSTS50011") {
        return "AADSTS50011 - redirect URI mismatch. The app registration has a reply URL, but not the http://localhost one this interactive flow uses. Add it under Authentication > Mobile and desktop applications."
    }
    if ($text -match "AADSTS700016" -or $text -match "AADSTS70001") {
        return "AADSTS700016 - no application with that Client ID exists in this tenant. Check the Client ID, and check that the Tenant ID matches the tenant the app is registered in."
    }
    if ($text -match "AADSTS90002") {
        return "AADSTS90002 - the Tenant ID was not found. Check it, or clear the field to sign in against your home tenant."
    }
    if ($text -match "AADSTS65001" -or $text -match "consent") {
        return "AADSTS65001 - the app has no consent for the permissions requested. Grant admin consent to the app registration in Entra, or clear the 'Do not request new permissions' checkbox to be prompted to consent - which permanently adds those permissions to this app registration."
    }
    if ($text -match "AADSTS50076" -or $text -match "AADSTS50079" -or $text -match "AADSTS53003") {
        return "Sign-in was blocked by a Conditional Access or MFA policy for this application. The account may be permitted to use the Intune portal while this app registration is out of policy scope."
    }

    # Checked last, so a specific AADSTS code always wins. This case matters more
    # than it looks: when Entra rejects the request in the browser, MSAL reports
    # only that the user cancelled, and the AADSTS code never reaches PowerShell.
    # An apparent cancellation is therefore often a registration fault.
    if ($text -match "canceled authentication" -or $text -match "cancelled authentication" -or $text -match "User canceled") {
        return "Sign-in was cancelled, or the browser window was closed before it completed.

If the browser displayed an AADSTS error rather than a sign-in page, that is the real failure - MSAL reports it here only as a cancellation, so the code stays in the browser window. The usual causes when using a custom app registration are:

    AADSTS500113  - no reply URL. Add http://localhost under
                    Authentication > Mobile and desktop applications.
    AADSTS7000218 - public client flows disabled. Set 'Allow public
                    client flows' to Yes under Authentication.
    AADSTS650057  - the app has not been granted the Microsoft Graph
                    permissions this tool needs."
    }

    return $null
}

# Translate a Graph failure into an actionable hint. Graph's own errors describe
# the transport, not what the operator needs to change.
function Get-GraphErrorHint {
    param ($ErrorRecord)

    $text = "$ErrorRecord"
    if ($text -match "403" -or $text -match "Forbidden" -or $text -match "not authorized") {
        return "403 Forbidden. The token was valid but the identity is not permitted to perform this write. Check, in order: (1) the app registration in use has the required delegated scopes consented - the built-in Microsoft Graph PowerShell app frequently cannot write to Intune even when it can read; (2) the signed-in account holds an Intune RBAC role granting update rights; (3) the account's scope tags include this object."
    }
    if ($text -match "401" -or $text -match "Unauthorized") {
        return "401 Unauthorized. The token was rejected or has expired. Re-authenticate on the Connect tab."
    }
    if ($text -match "429" -or $text -match "TooManyRequests") {
        return "429 Throttled. Graph is rate limiting this tenant. Retry with fewer objects selected."
    }
    return $null
}

# Issue a GET and follow @odata.nextLink, returning every value entry.
function Get-GraphCollection {
    param (
        [string]$Uri,
        [hashtable]$Headers
    )

    $results = New-Object System.Collections.Generic.List[object]
    $next    = $Uri
    do {
        if ($Headers) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers -ErrorAction Stop
        } else {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        }
        foreach ($item in $resp.value) { $results.Add($item) }
        $next = $resp.'@odata.nextLink'
    } while ($next)

    return $results
}

# Retrieve a collection with an optional display-name prefix filter.
#
# Not every Intune endpoint allows startswith() in $filter. deviceCompliancePolicies
# in particular rejects it with:
#   "Function 'startswith' is not allowed. To allow it, set the 'AllowedFunctions'
#    property on EnableQueryAttribute or QueryValidationSettings."
# Rather than maintain a per-endpoint allow list that will drift as the service
# changes, this tries the server-side filter first and transparently falls back to
# retrieving the collection and matching the prefix client-side.
#
# BaseUri must already contain a query string (for example "...?`$select=id,displayName")
# so the filter can be appended with "&".
function Get-GraphCollectionByPrefix {
    param (
        [string]$BaseUri,
        [string]$NameProperty,   # displayName | name
        [string]$Prefix
    )

    if ([string]::IsNullOrWhiteSpace($Prefix)) {
        return Get-GraphCollection -Uri $BaseUri
    }

    try {
        $filtered = "$BaseUri&`$filter=startswith($NameProperty,'$(ConvertTo-ODataLiteral $Prefix)')"
        return Get-GraphCollection -Uri $filtered
    } catch {
        Write-LogMessage "Endpoint rejected a server-side startswith() filter; retrieving the full collection and filtering on '$Prefix' client-side. Uri: $BaseUri" "WARN"
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in (Get-GraphCollection -Uri $BaseUri)) {
        $name = "$($item.$NameProperty)"
        if ($name.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $results.Add($item)
        }
    }
    return $results
}

# Execute many GETs through the Graph JSON batch endpoint.
#
# Sequential per-object GETs are correct but slow: several hundred round trips to
# survey a tenant. $expand=assignments would collapse that into one call, but it
# is deprecated on mobileApps and unreliable on the configuration profile
# endpoints, which return 503 and 504 under it. An intermittently short result is
# the worst possible input to a tool that removes assignments, so this batches
# individual requests instead - the same approach the Intune portal itself uses.
#
# $Requests is a list of hashtables with Key and Url (Url relative to the beta
# root, e.g. "/deviceManagement/configurationPolicies('id')/assignments").
# Returns a hashtable of Key -> @{ Status; Value; Error }.
function Invoke-GraphBatchGet {
    param (
        [object[]]$Requests,
        [int]$BatchSize = 20,
        [int]$MaxAttempts = 4
    )

    $results = @{}
    $pending = New-Object System.Collections.Generic.List[object]
    foreach ($request in $Requests) { $pending.Add($request) }

    $requestedKeys = @($Requests | ForEach-Object { $_.Key })

    for ($attempt = 1; $attempt -le $MaxAttempts -and $pending.Count -gt 0; $attempt++) {
        $retry         = New-Object System.Collections.Generic.List[object]
        $attemptOk     = 0
        $attemptFailed = 0
        $attemptStart  = [System.Diagnostics.Stopwatch]::StartNew()

        for ($offset = 0; $offset -lt $pending.Count; $offset += $BatchSize) {
            $chunk = @($pending[$offset..([Math]::Min($offset + $BatchSize - 1, $pending.Count - 1))])

            # The batch response is not guaranteed to preserve request order, so
            # each request carries an index that maps back to its Key.
            $indexToKey  = @{}
            $batchItems  = New-Object System.Collections.Generic.List[object]
            for ($i = 0; $i -lt $chunk.Count; $i++) {
                $indexToKey["$i"] = $chunk[$i].Key
                $batchItems.Add(@{ id = "$i"; method = "GET"; url = $chunk[$i].Url })
            }

            $body = @{ requests = $batchItems.ToArray() } | ConvertTo-Json -Depth 10

            try {
                $response = Invoke-MgGraphRequest -Method POST -Uri "$graphBeta/`$batch" -Body $body -ContentType "application/json" -ErrorAction Stop
            } catch {
                # A failure of the batch envelope itself, not of the individual
                # requests. Retry the whole chunk.
                Write-LogMessage "Batch request failed (attempt $attempt): $_" "WARN"
                foreach ($item in $chunk) { $retry.Add($item) }
                continue
            }

            foreach ($item in @($response.responses)) {
                $key    = $indexToKey["$($item.id)"]
                $status = [int]$item.status

                if ($status -ge 200 -and $status -lt 300) {
                    $results[$key] = @{ Status = $status; Value = @($item.body.value); Error = $null }
                    $attemptOk++
                } elseif ($status -eq 429 -or $status -eq 503 -or $status -eq 504) {
                    # Throttling or a transient service fault - retry this one.
                    $original = @($chunk | Where-Object { $_.Key -eq $key })
                    if ($original.Count -gt 0) { $retry.Add($original[0]) }
                } else {
                    # A real failure for this object (403 outside scope tags, 404
                    # deleted since listing). Recorded, not retried, never silently
                    # dropped.
                    $message = "$($item.body.error.code): $($item.body.error.message)"
                    $results[$key] = @{ Status = $status; Value = @(); Error = $message }
                    $attemptFailed++
                }
            }
        }

        $attemptStart.Stop()

        # Record every attempt, not just the fact that a retry was scheduled.
        # Without this a retry that succeeds is indistinguishable in the log from
        # one that was quietly dropped - the totals only line up afterwards, and
        # working that out by arithmetic is not a reasonable thing to ask.
        Write-LogMessage ("Batch attempt {0}: {1} request(s) sent, {2} succeeded, {3} failed outright, {4} queued for retry, {5:N1}s." -f `
            $attempt, $pending.Count, $attemptOk, $attemptFailed, $retry.Count, $attemptStart.Elapsed.TotalSeconds)

        $pending = $retry
        if ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts) {
            $delay = [Math]::Pow(2, $attempt)
            Write-LogMessage "$($pending.Count) batched request(s) were throttled or failed transiently; retrying in $delay second(s)."
            Start-Sleep -Seconds $delay
        }
    }

    # Anything still pending exhausted its retries. Record it as an error rather
    # than leaving the caller to treat a missing key as "no assignments".
    foreach ($item in $pending) {
        $results[$item.Key] = @{ Status = 0; Value = @(); Error = "Exhausted $MaxAttempts attempts (throttled or transient service errors)." }
    }

    # Every requested key must come back with an entry. A missing key would be
    # read by the caller as "no assignments" rather than "not answered", which is
    # the one failure mode this whole approach exists to avoid - so if the
    # accounting does not balance, say so loudly and mark the gaps as errors.
    $missingKeys = @($requestedKeys | Where-Object { -not $results.ContainsKey($_) })
    if ($missingKeys.Count -gt 0) {
        Write-LogMessage "Batch accounting mismatch: $($missingKeys.Count) of $($requestedKeys.Count) request(s) received no response at all. Treating them as unreadable." "ERROR"
        foreach ($key in $missingKeys) {
            $results[$key] = @{ Status = 0; Value = @(); Error = "No response was returned for this request." }
        }
    }

    Write-LogMessage "Batch complete: $($results.Count) result(s) for $($requestedKeys.Count) request(s); $(@($results.Values | Where-Object { $_.Error }).Count) unreadable."
    return $results
}

# Build the relative batch URL for an object's assignments collection.
function Get-AssignmentBatchUrl {
    param ($Object)

    if ($Object.Kind -eq "mobileApps") {
        return "/deviceAppManagement/mobileApps/$($Object.Id)/assignments"
    }
    if ($Object.Kind -eq "mobileAppConfigurations") {
        return "/deviceAppManagement/mobileAppConfigurations/$($Object.Id)/assignments"
    }
    return "/deviceManagement/$($Object.Kind)('$($Object.Id)')/assignments"
}

# Retrieve assignments for many objects at once. Returns a hashtable of object
# Id -> @{ Status; Value; Error }.
function Get-AssignmentMap {
    param ([object[]]$Objects)

    $requests = New-Object System.Collections.Generic.List[object]
    foreach ($object in $Objects) {
        $requests.Add(@{ Key = "$($object.Id)"; Url = (Get-AssignmentBatchUrl -Object $object) })
    }

    Write-LogMessage "Reading assignments for $($requests.Count) object(s) via batched requests."
    return Invoke-GraphBatchGet -Requests $requests.ToArray()
}

# Escape a single quote for use inside an OData string literal.
function ConvertTo-ODataLiteral {
    param ([string]$Value)
    if ($null -eq $Value) { return "" }
    return $Value.Replace("'", "''")
}

# ---------------------------------------------------------------------------
# Helper functions - object retrieval
# ---------------------------------------------------------------------------

# Resolve the platform from a Graph @odata.type, returning the label used by the
# Objects grid. Prefix matching avoids false positives such as "ios" inside a
# Windows type name.
function Get-PlatformFromOdataType {
    param ([string]$OdataType)

    if ([string]::IsNullOrWhiteSpace($OdataType)) { return $null }
    $name = $OdataType -replace '^#microsoft\.graph\.', ''
    if ($name -match '^(ios|iphone)') { return 'iOS' }
    if ($name -match '^(managed)?android') { return 'Android' }
    if ($name -match '^macos') { return 'macOS' }
    if ($name -match '^(windows|win32|microsoftstore)' -or $name -eq 'officesuiteapp') { return 'Windows' }
    if ($global:WindowsExtraOdataTypes -contains $name) { return 'Windows' }
    return $null
}

# Resolve the platform value returned by configurationPolicies.
function Get-PlatformFromPolicyValue {
    param ([object]$Platforms)

    $value = (@($Platforms) -join ' ').ToLowerInvariant()
    if ($value -match 'ios|iphone') { return 'iOS' }
    if ($value -match 'android') { return 'Android' }
    if ($value -match 'macos|mac') { return 'macOS' }
    if ($value -match 'windows') { return 'Windows' }
    return $null
}

# Retrieve supported mobile apps (optionally filtered by display-name prefix).
# Returns DisplayName, Id, Kind, OdataType, Platform, and TypeLabel.
function Get-MobileApps {
    param ([string]$Prefix)

    $uri = "$graphBeta/deviceAppManagement/mobileApps?`$select=id,displayName,roleScopeTagIds"
    $raw = Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "displayName" -Prefix $Prefix

    # Filter supported app types client-side because the mobileApps endpoint
    # does not reliably support combining isof() with a startswith filter.
    $results = New-Object System.Collections.Generic.List[object]
    $skippedTypes = New-Object System.Collections.Generic.List[string]
    foreach ($item in $raw) {
        $odataType = $item.'@odata.type'
        $platform = Get-PlatformFromOdataType $odataType
        if (-not $platform) {
            if ($skippedTypes -notcontains $odataType) { $skippedTypes.Add($odataType) }
            continue
        }
        $results.Add([pscustomobject]@{
            DisplayName     = $item.displayName
            Id              = $item.id
            Kind            = "mobileApps"
            OdataType       = $odataType
            Platform        = $platform
            TypeLabel       = ($odataType -replace '^#microsoft\.graph\.', '')
            RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
        })
    }

    if ($skippedTypes.Count -gt 0) {
        Write-LogMessage "Skipped app types with an unsupported platform: $($skippedTypes -join ', ')"
    }

    return $results | Sort-Object DisplayName
}

# Retrieve supported configuration profiles across all relevant Graph object types,
# including remediation scripts.
# Returns DisplayName, Id, Kind, OdataType, Platform, and TypeLabel.
function Get-ConfigurationProfiles {
    param ([string]$Prefix)

    $results      = New-Object System.Collections.Generic.List[object]
    $skippedTypes = New-Object System.Collections.Generic.List[string]

    # --- 1. App configuration policies -----------------------------------
    $uri = "$graphBeta/deviceAppManagement/mobileAppConfigurations?`$select=id,displayName,roleScopeTagIds"
    try {
        foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "displayName" -Prefix $Prefix)) {
            $odataType = $item.'@odata.type'
            $platform = Get-PlatformFromOdataType $odataType
            if (-not $platform -and $odataType -eq "#microsoft.graph.targetedManagedAppConfiguration") {
                $platform = "Other"
            }
            if (-not $platform) {
                if ($skippedTypes -notcontains $odataType) { $skippedTypes.Add($odataType) }
                continue
            }
            $results.Add([pscustomobject]@{
                DisplayName     = $item.displayName
                Id              = $item.id
                Kind            = "mobileAppConfigurations"
                OdataType       = $odataType
                Platform        = $platform
                TypeLabel       = "App Configuration"
                RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
            })
        }
    } catch {
        Write-LogMessage "Error retrieving mobileAppConfigurations: $_" "ERROR"
        throw
    }

    # --- 2. Settings catalog and templates -------------------------------
    # Note: the display-name property here is "name", not "displayName".
    $uri = "$graphBeta/deviceManagement/configurationPolicies?`$select=id,name,platforms,technologies,roleScopeTagIds"
    try {
        foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "name" -Prefix $Prefix)) {
            $platform = Get-PlatformFromPolicyValue $item.platforms
            if (-not $platform) { continue }
            $results.Add([pscustomobject]@{
                DisplayName     = $item.name
                Id              = $item.id
                Kind            = "configurationPolicies"
                OdataType       = $null      # configurationPolicy has no @odata.type
                Platform        = $platform
                TypeLabel       = "Settings Catalog"
                RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
            })
        }
    } catch {
        Write-LogMessage "Error retrieving configurationPolicies: $_" "ERROR"
        throw
    }

    # --- 3. Legacy device configuration profiles -------------------------
    $uri = "$graphBeta/deviceManagement/deviceConfigurations?`$select=id,displayName,roleScopeTagIds"
    try {
        foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "displayName" -Prefix $Prefix)) {
            $odataType = $item.'@odata.type'
            $platform = Get-PlatformFromOdataType $odataType
            if (-not $platform) {
                if ($skippedTypes -notcontains $odataType) { $skippedTypes.Add($odataType) }
                continue
            }
            $results.Add([pscustomobject]@{
                DisplayName     = $item.displayName
                Id              = $item.id
                Kind            = "deviceConfigurations"
                OdataType       = $odataType
                Platform        = $platform
                TypeLabel       = "Device Configuration"
                RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
            })
        }
    } catch {
        Write-LogMessage "Error retrieving deviceConfigurations: $_" "ERROR"
        throw
    }

    # --- 4. ADMX / administrative templates ------------------------------
    # groupPolicyConfigurations are Windows-only by definition.
    $uri = "$graphBeta/deviceManagement/groupPolicyConfigurations?`$select=id,displayName,roleScopeTagIds"
    try {
        foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "displayName" -Prefix $Prefix)) {
            $results.Add([pscustomobject]@{
                DisplayName     = $item.displayName
                Id              = $item.id
                Kind            = "groupPolicyConfigurations"
                OdataType       = $null
                Platform        = "Windows"
                TypeLabel       = "ADMX Template"
                RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
            })
        }
    } catch {
        Write-LogMessage "Error retrieving groupPolicyConfigurations: $_" "ERROR"
        throw
    }

    # --- 5. Compliance policies ------------------------------------------
    $uri = "$graphBeta/deviceManagement/deviceCompliancePolicies?`$select=id,displayName,roleScopeTagIds"
    try {
        foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $uri -NameProperty "displayName" -Prefix $Prefix)) {
            $odataType = $item.'@odata.type'
            $platform = Get-PlatformFromOdataType $odataType
            if (-not $platform) {
                if ($skippedTypes -notcontains $odataType) { $skippedTypes.Add($odataType) }
                continue
            }
            $results.Add([pscustomobject]@{
                DisplayName     = $item.displayName
                Id              = $item.id
                Kind            = "deviceCompliancePolicies"
                OdataType       = $odataType
                Platform        = $platform
                TypeLabel       = "Compliance Policy"
                RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
            })
        }
    } catch {
        Write-LogMessage "Error retrieving deviceCompliancePolicies: $_" "ERROR"
        throw
    }

    # --- 6. Platform scripts and remediation scripts -------------------
    # Intune exposes several script families depending on platform and tenant
    # capabilities. Include both Windows and macOS script variants so the object
    # list reflects what is actually available in Graph.
    $scriptEndpoints = @(
        @{ Uri = "$graphBeta/deviceManagement/deviceManagementScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "deviceManagementScripts"; TypeLabel = "Platform Script"; Platform = "Windows" },
        @{ Uri = "$graphBeta/deviceManagement/managedDeviceScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "managedDeviceScripts"; TypeLabel = "Remediation Script"; Platform = "Windows" },
        @{ Uri = "$graphBeta/deviceManagement/deviceHealthScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "deviceHealthScripts"; TypeLabel = "Remediation Script"; Platform = "Windows" },
        @{ Uri = "$graphBeta/deviceManagement/managedDeviceHealthScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "managedDeviceHealthScripts"; TypeLabel = "Remediation Script"; Platform = "Windows" },
        @{ Uri = "$graphBeta/deviceManagement/deviceShellScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "deviceShellScripts"; TypeLabel = "Shell Script"; Platform = "macOS" },
        @{ Uri = "$graphBeta/deviceManagement/managedDeviceShellScripts?`$select=id,displayName,roleScopeTagIds"; Kind = "managedDeviceShellScripts"; TypeLabel = "Shell Script"; Platform = "macOS" },
        @{ Uri = "$graphBeta/deviceManagement/deviceShellScripts?`$filter=platform eq 'macOS'&`$select=id,displayName,roleScopeTagIds"; Kind = "deviceShellScriptsMac"; TypeLabel = "macOS Script"; Platform = "macOS" },
        @{ Uri = "$graphBeta/deviceManagement/managedDeviceShellScripts?`$filter=platform eq 'macOS'&`$select=id,displayName,roleScopeTagIds"; Kind = "managedDeviceShellScriptsMac"; TypeLabel = "macOS Script"; Platform = "macOS" }
    )

    foreach ($scriptDef in $scriptEndpoints) {
        try {
            foreach ($item in (Get-GraphCollectionByPrefix -BaseUri $scriptDef.Uri -NameProperty "displayName" -Prefix $Prefix)) {
                $key = "$($scriptDef.Kind):$($item.id)"
                if ($results | Where-Object { "$($_.Kind):$($_.Id)" -eq $key }) { continue }

                $results.Add([pscustomobject]@{
                    DisplayName     = $item.displayName
                    Id              = $item.id
                    Kind            = $scriptDef.Kind
                    OdataType       = $null
                    Platform        = $scriptDef.Platform
                    TypeLabel       = $scriptDef.TypeLabel
                    RoleScopeTagIds = @($item.roleScopeTagIds | ForEach-Object { "$_" })
                })
            }
        } catch {
            Write-LogMessage "Script endpoint not available or not supported: $($scriptDef.Uri). Details: $_" "WARN"
        }
    }

    if ($skippedTypes.Count -gt 0) {
        Write-LogMessage "Skipped profile types with an unsupported platform: $($skippedTypes -join ', ')"
    }

    return $results | Sort-Object DisplayName
}

# Return every supported object in scope for the current mode, with no name
# filter. Used by Remove discovery, which must survey the whole estate.
function Get-AllSupportedObjects {
    param ([string]$Mode)

    $results = New-Object System.Collections.Generic.List[object]
    if ($Mode -eq "Apps" -or $Mode -eq "Both") {
        foreach ($item in (Get-MobileApps -Prefix "")) { $results.Add($item) }
    }
    if ($Mode -eq "Profiles" -or $Mode -eq "Both") {
        foreach ($item in (Get-ConfigurationProfiles -Prefix "")) { $results.Add($item) }
    }
    return $results
}

# Find every supported object that targets the given group.
#
# SCALE: there is no server-side way to ask Graph which objects target a group.
# The only way to know is to read every object's assignments, so the cost grows
# linearly with the size of the estate - a tenant with a few thousand apps and
# profiles means a few thousand assignment reads. Batching cuts that to one
# request per twenty objects, but a large estate is still a wait, and "Both"
# mode is the slowest because it surveys apps and all four profile types.
#
# Each returned object carries the same fields the Objects grid already uses,
# plus a description of how the group is currently targeted. Two of those matter
# more than they look:
#
#   Exclusions. Removing an EXCLUSION widens a policy's scope rather than
#   narrowing it. That is the opposite of what "remove a group" sounds like, so
#   the assignment type is shown rather than left implicit.
#
#   Sole assignments. If the group is the object's only assignment, removing it
#   leaves a policy assigned to nobody - still present, still maintained, and
#   silently applying to nothing. Flagged so it is a decision, not a surprise.
function Find-ObjectsAssignedToGroup {
    param (
        [string]$GroupId,
        [object[]]$Objects
    )

    # The object list is passed in rather than fetched here so the caller can
    # size the survey and warn before the expensive part begins.
    $allObjects = @($Objects)
    Write-LogMessage "Surveying $($allObjects.Count) object(s) for assignments to group $GroupId."

    $surveyTimer   = [System.Diagnostics.Stopwatch]::StartNew()
    $assignmentMap = Get-AssignmentMap -Objects $allObjects
    $surveyTimer.Stop()
    Write-LogMessage "Assignment survey of $($allObjects.Count) object(s) took $([Math]::Round($surveyTimer.Elapsed.TotalSeconds, 1)) second(s)."

    $matchedObjects = New-Object System.Collections.Generic.List[object]
    $unreadable     = New-Object System.Collections.Generic.List[string]

    foreach ($object in $allObjects) {
        $entry = $assignmentMap["$($object.Id)"]

        # An object whose assignments could not be read is NOT evidence that the
        # group is absent. Surfacing it separately keeps a transient failure from
        # quietly narrowing the removal set.
        if ($null -eq $entry -or $entry.Error) {
            $reason = "could not read assignments"
            if ($entry -and $entry.Error) { $reason = $entry.Error }
            $unreadable.Add("$($object.DisplayName) [$($object.TypeLabel)] - $reason")
            continue
        }

        $existing = @($entry.Value)
        $hits     = @($existing | Where-Object { $_.target.groupId -eq $GroupId })
        if ($hits.Count -eq 0) { continue }

        $descriptions = New-Object System.Collections.Generic.List[string]
        foreach ($hit in $hits) {
            if ($hit.target.'@odata.type' -like "*exclusionGroupAssignmentTarget*") {
                $kind = "Exclude"
            } else {
                $kind = "Include"
            }
            $filterType = "$($hit.target.deviceAndAppManagementAssignmentFilterType)"
            if ($filterType -eq "include" -or $filterType -eq "exclude") {
                $kind += " (filter: $filterType)"
            }
            $descriptions.Add($kind)
        }

        $summary = ($descriptions -join ", ")
        if ($existing.Count -eq $hits.Count) { $summary += "  - ONLY ASSIGNMENT" }

        $matchedObjects.Add([pscustomobject]@{
            DisplayName        = $object.DisplayName
            Id                 = $object.Id
            Kind               = $object.Kind
            OdataType          = $object.OdataType
            Platform           = $object.Platform
            TypeLabel          = $object.TypeLabel
            AssignmentSummary  = $summary
            ExistingCount      = $existing.Count
            MatchCount         = $hits.Count
        })
    }

    Write-LogMessage "$($matchedObjects.Count) object(s) target group $GroupId. $($unreadable.Count) object(s) could not be read."
    foreach ($line in $unreadable) { Write-LogMessage "Unreadable during discovery: $line" "WARN" }

    return [pscustomobject]@{
        Matches    = $matchedObjects.ToArray()
        Unreadable = $unreadable.ToArray()
        Surveyed   = $allObjects.Count
    }
}

# Find every supported object carrying any of the given scope tags.
#
# Cheap, unlike its group-assignment counterpart. roleScopeTagIds is a plain
# property on all five object types rather than a relationship, so it comes back
# with the list call itself - no per-object reads, no batching, no survey. The
# same question about group assignments costs one request per object because
# assignments are a relationship and $expand over them is unusable.
function Find-ObjectsWithScopeTag {
    param (
        [string[]]$TagIds,
        [string]$Mode
    )

    $allObjects = @(Get-AllSupportedObjects -Mode $Mode)
    Write-LogMessage "Checking $($allObjects.Count) object(s) for scope tag(s) $($TagIds -join ', ')."

    $matchedObjects = New-Object System.Collections.Generic.List[object]
    foreach ($object in $allObjects) {
        $objectTags = @($object.RoleScopeTagIds)
        $hits       = @($objectTags | Where-Object { $TagIds -contains $_ })
        if ($hits.Count -eq 0) { continue }

        # Removing every tag an object has would leave it invisible to scoped
        # administrators, so the same last-tag rule that governs the write is
        # surfaced here, before anything is selected.
        $summary = "Has: " + ($hits -join ", ")
        if ($objectTags.Count -eq $hits.Count) { $summary += "  - ONLY SCOPE TAG(S)" }

        $matchedObjects.Add([pscustomobject]@{
            DisplayName       = $object.DisplayName
            Id                = $object.Id
            Kind              = $object.Kind
            OdataType         = $object.OdataType
            Platform          = $object.Platform
            TypeLabel         = $object.TypeLabel
            RoleScopeTagIds   = $objectTags
            AssignmentSummary = $summary
            ExistingCount     = $objectTags.Count
            MatchCount        = $hits.Count
        })
    }

    Write-LogMessage "$($matchedObjects.Count) object(s) carry the selected scope tag(s)."
    return [pscustomobject]@{
        Matches    = $matchedObjects.ToArray()
        Unreadable = @()
        Surveyed   = $allObjects.Count
    }
}

# Search Entra ID for groups by display name. Deliberately search-based rather
# than a bulk enumeration of every group in the tenant.
#
# Prefers $search (substring match anywhere in the name), which requires the
# ConsistencyLevel: eventual header. Falls back to startswith() if the tenant or
# session rejects the advanced query.
function Search-EntraGroups {
    param (
        [string]$SearchText,
        [int]$Top = 100
    )

    if ([string]::IsNullOrWhiteSpace($SearchText)) {
        throw "Enter at least one character to search for."
    }

    $clean = $SearchText.Replace('"', '').Trim()

    try {
        $uri = "$graphV1/groups?`$search=`"displayName:$clean`"&`$select=id,displayName&`$top=$Top&`$count=true"
        $raw = Get-GraphCollection -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' }
    } catch {
        Write-LogMessage "Advanced group search failed ($_); falling back to startswith()." "WARN"
        $uri = "$graphV1/groups?`$filter=startswith(displayName,'$(ConvertTo-ODataLiteral $clean)')&`$select=id,displayName&`$top=$Top"
        $raw = Get-GraphCollection -Uri $uri
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $raw) {
        $results.Add([pscustomobject]@{ DisplayName = $item.displayName; Id = $item.id })
    }
    return $results | Sort-Object DisplayName
}

# Retrieve Windows device assignment filters. App-management filters
# (assignmentFilterManagementType 'apps') are for managed-app / MAM scenarios
# and are not valid on Win32 app or configuration profile assignments, so they
# are excluded here.
function Get-WindowsAssignmentFilters {
    $uri = "$graphBeta/deviceManagement/assignmentFilters?`$filter=platform eq 'windows10AndLater'"
    $raw = Get-GraphCollection -Uri $uri

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $raw) {
        $mgmtType = $item.assignmentFilterManagementType
        if ($mgmtType -and $mgmtType -ne 'devices') { continue }
        $results.Add([pscustomobject]@{ DisplayName = $item.displayName; Id = $item.id })
    }
    return $results | Sort-Object DisplayName
}

# Retrieve every role scope tag. The built-in Default tag is ID "0".
function Get-RoleScopeTags {
    $uri = "$graphBeta/deviceManagement/roleScopeTags?`$select=id,displayName"
    $raw = Get-GraphCollection -Uri $uri

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $raw) {
        $results.Add([pscustomobject]@{ DisplayName = $item.displayName; Id = "$($item.id)" })
    }
    return $results | Sort-Object { [int]$_.Id }
}

# ---------------------------------------------------------------------------
# Helper functions - assignment targets
# ---------------------------------------------------------------------------

# Build an assignment target hashtable for a group, including filter properties.
function New-AssignmentTarget {
    param (
        [string]$GroupId,
        [string]$Type,        # Include | Exclude
        [string]$FilterId,
        [string]$FilterMode   # none | include | exclude
    )

    if ($Type -eq "Exclude") {
        $targetType = "#microsoft.graph.exclusionGroupAssignmentTarget"
    } else {
        $targetType = "#microsoft.graph.groupAssignmentTarget"
    }

    $target = @{
        "@odata.type" = $targetType
        groupId       = $GroupId
    }

    # Intune expects the filter type to always be present. "none" with a null ID
    # is the explicit "no filter" state.
    if ($FilterMode -eq 'include' -or $FilterMode -eq 'exclude') {
        $target.deviceAndAppManagementAssignmentFilterId   = $FilterId
        $target.deviceAndAppManagementAssignmentFilterType = $FilterMode
    } else {
        $target.deviceAndAppManagementAssignmentFilterId   = $null
        $target.deviceAndAppManagementAssignmentFilterType = "none"
    }

    return $target
}

# Base URI for a configuration profile object, using the parenthesised key form
# that the Intune portal itself uses against these endpoints.
function Get-ProfileUri {
    param (
        [string]$Kind,
        [string]$Id
    )
    if ($Kind -eq "mobileAppConfigurations") {
        return "$graphBeta/deviceAppManagement/mobileAppConfigurations('$Id')"
    }
    return "$graphBeta/deviceManagement/$Kind('$Id')"
}

# ---------------------------------------------------------------------------
# Helper functions - app assignments (true additive collection)
# ---------------------------------------------------------------------------

# Add an Available group-assignment to a single app (Include or Exclude), with
# an optional assignment filter. Idempotent - returns a status string:
#   "Added"         - assignment created
#   "AlreadyExists" - the group is already assigned with the same target type
#   "Conflict"      - the group is already assigned with the OPPOSITE type
#                     (Intune does not allow both include and exclude for one
#                      group on one app), so the add is skipped, not errored
#
# The settings block is tailored to the app type. Win32 apps get
# win32LobAppAssignmentSettings; restart and install-time settings are only
# valid for "required" intent, so they are included only when Intent is
# "required".
function Add-AppAssignment {
    param (
        [string]$AppId,
        [string]$AppType,
        [hashtable]$Target,
        [string]$Intent = "available"
    )

    $uri = "$graphBeta/deviceAppManagement/mobileApps/$AppId/assignments"

    # Pre-check: is this group already assigned to the app?
    $existing = @((Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop).value)
    $match    = $existing | Where-Object { $_.target.groupId -eq $Target.groupId } | Select-Object -First 1
    if ($match) {
        if ($match.target.'@odata.type' -eq $Target.'@odata.type') {
            return "AlreadyExists"
        }
        return "Conflict"
    }

    $body = @{
        target = $Target
        intent = $Intent
    }

    if ($global:Win32AppTypes -contains $AppType) {
        $settings = @{
            "@odata.type"                = "#microsoft.graph.win32LobAppAssignmentSettings"
            notifications                = "showReboot"
            deliveryOptimizationPriority = "notConfigured"
        }
        if ($Intent -eq "required") {
            $settings.installTimeSettings = $null
            $settings.restartSettings = @{
                "@odata.type"                              = "#microsoft.graph.win32LobAppRestartSettings"
                gracePeriodInMinutes                       = 720
                countdownDisplayBeforeRestartInMinutes     = 90
                restartNotificationSnoozeDurationInMinutes = 15
            }
        }
        $body.settings = $settings
    }

    $json = $body | ConvertTo-Json -Depth 10
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop | Out-Null
    return "Added"
}

# Remove every assignment on an app that targets the given group (group match
# only). Returns the number of assignments removed.
function Remove-AppAssignment {
    param (
        [string]$AppId,
        [string]$GroupId
    )

    $uri         = "$graphBeta/deviceAppManagement/mobileApps/$AppId/assignments"
    $assignments = @((Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop).value)
    $matching    = @($assignments | Where-Object { $_.target.groupId -eq $GroupId })

    $removed = 0
    foreach ($assignment in $matching) {
        $delUri = "$uri/$($assignment.id)"
        Invoke-MgGraphRequest -Method DELETE -Uri $delUri -ErrorAction Stop | Out-Null
        $removed++
    }
    return $removed
}

# ---------------------------------------------------------------------------
# Helper functions - profile assignments (read-merge-write)
# ---------------------------------------------------------------------------

# Read the current assignment set for a configuration profile. Throws on
# failure so the caller never proceeds to a destructive /assign with an
# incomplete picture.
function Get-ProfileAssignment {
    param (
        [string]$Kind,
        [string]$Id
    )
    $uri = "$(Get-ProfileUri -Kind $Kind -Id $Id)/assignments"
    $assignments = @((Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop).value)

    # The leading comma is load-bearing. PowerShell unrolls a returned array, so
    # a profile with exactly one assignment would hand the caller a bare
    # Hashtable instead of a one-element array - and $existing.Count would then
    # return the Hashtable's KEY count (id, target, source, sourceId = 4) rather
    # than 1. Wrapping the array keeps it intact across the return boundary.
    return ,$assignments
}

# Normalise an assignment target into a payload Graph will accept on /assign.
#
# A target read back from GET .../assignments is NOT automatically valid to POST
# again. The read model carries properties that the write model rejects - the
# combination varies by target type and by how the assignment was originally
# created (portal, Graph, or a policy migration) - and the result is a
# 400 "Property target in payload has a value that does not match schema" that
# names no property and no assignment.
#
# Rather than echo back whatever was read, each target is rebuilt from an
# allow list of the properties that are valid for its own @odata.type.
function ConvertTo-AssignmentTargetPayload {
    param ($Target)

    $odataType = "$($Target.'@odata.type')"

    # An existing target should always declare its type, but a missing one is
    # itself a cause of this exact 400, so infer it rather than pass it through.
    if ([string]::IsNullOrWhiteSpace($odataType)) {
        if ($Target.collectionId) {
            $odataType = "#microsoft.graph.configurationManagerCollectionAssignmentTarget"
        } elseif ($Target.groupId) {
            $odataType = "#microsoft.graph.groupAssignmentTarget"
        } else {
            Write-LogMessage "An existing assignment target has no @odata.type and no group or collection id; passing it through unchanged." "WARN"
            return $Target
        }
        Write-LogMessage "An existing assignment target had no @odata.type; inferred '$odataType'." "WARN"
    }

    $clean = @{ "@odata.type" = $odataType }

    # Only the identifier that belongs to this target type. Carrying a null
    # groupId onto an all-devices target, for example, fails schema validation.
    switch ($odataType) {
        "#microsoft.graph.groupAssignmentTarget"          { $clean.groupId      = "$($Target.groupId)" }
        "#microsoft.graph.exclusionGroupAssignmentTarget" { $clean.groupId      = "$($Target.groupId)" }
        "#microsoft.graph.configurationManagerCollectionAssignmentTarget" { $clean.collectionId = "$($Target.collectionId)" }
        "#microsoft.graph.allDevicesAssignmentTarget"     { }
        "#microsoft.graph.allLicensedUsersAssignmentTarget" { }
        default {
            # An unrecognised target type is preserved as-is rather than
            # stripped, so an unfamiliar assignment is never silently dropped
            # from the merged set.
            Write-LogMessage "Unrecognised assignment target type '$odataType'; preserving it unchanged." "WARN"
            return $Target
        }
    }

    # Filters are valid on every target type above, but the id must be absent
    # unless the type is include or exclude.
    $filterType = "$($Target.deviceAndAppManagementAssignmentFilterType)"
    if ($filterType -eq "include" -or $filterType -eq "exclude") {
        $clean.deviceAndAppManagementAssignmentFilterId   = "$($Target.deviceAndAppManagementAssignmentFilterId)"
        $clean.deviceAndAppManagementAssignmentFilterType = $filterType
    } else {
        $clean.deviceAndAppManagementAssignmentFilterId   = $null
        $clean.deviceAndAppManagementAssignmentFilterType = "none"
    }

    return $clean
}

# Write a full assignment set back to a configuration profile.
#
# CAUTION: /assign replaces everything. Only ever call this with a complete,
# already-merged target list.
function Set-ProfileAssignment {
    param (
        [string]$Kind,
        [string]$Id,
        [object[]]$Targets
    )

    # Safety net. A nested collection here means a caller double-wrapped an array,
    # which would post several targets inside a single assignment entry and fail
    # schema validation. Flattening yields the correct target list either way, and
    # the warning makes the caller's mistake visible instead of silent.
    $flattened = New-Object System.Collections.Generic.List[object]
    foreach ($target in $Targets) {
        if ($target -is [System.Collections.IEnumerable] -and
            $target -isnot [System.Collections.IDictionary] -and
            $target -isnot [string]) {
            Write-LogMessage "An assignment target was nested inside a collection; flattening it. A caller passed a nested array." "WARN"
            foreach ($innerTarget in $target) { $flattened.Add($innerTarget) }
        } else {
            $flattened.Add($target)
        }
    }

    $assignments = New-Object System.Collections.Generic.List[object]
    foreach ($target in $flattened) {
        $assignments.Add(@{ target = (ConvertTo-AssignmentTargetPayload -Target $target) })
    }

    if ($Kind -eq "mobileAppConfigurations") {
        $body = @{ mobileAppAssignments = $assignments.ToArray() } | ConvertTo-Json -Depth 20
    } else {
        $body = @{ assignments = $assignments.ToArray() } | ConvertTo-Json -Depth 20
    }
    $uri  = "$(Get-ProfileUri -Kind $Kind -Id $Id)/assign"

    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
    } catch {
        # Graph's schema errors name neither the property nor the offending
        # assignment, so log the exact payload that was rejected. Without this
        # the only way to diagnose a bad target is to guess.
        Write-LogMessage "Rejected /assign payload for $Kind/$Id was: $body" "ERROR"
        throw
    }
}

# Additively add a target to a configuration profile's assignment set.
# Returns a hashtable: Status, Before, After.
#   Status "Added"         - target merged in and written back
#   Status "AlreadyExists" - the group is already assigned with the same type
#   Status "Conflict"      - the group is already assigned with the opposite type
function Add-ProfileAssignment {
    param (
        [string]$Kind,
        [string]$Id,
        [hashtable]$Target
    )

    # Get-ProfileAssignment already returns the array intact (see the comma in
    # its return statement). Do NOT wrap this in @() - collecting an output
    # stream whose single item is itself an array produces a NESTED array, and
    # the foreach below then iterates once with $assignment bound to the whole
    # array. $assignment.target would member-enumerate into an array of targets,
    # which Graph rejects with "Property target in payload has a value that does
    # not match schema". A profile with exactly one existing assignment happens
    # to survive that, which is what made this look intermittent.
    $existing = Get-ProfileAssignment -Kind $Kind -Id $Id
    $before   = $existing.Count

    foreach ($assignment in $existing) {
        if ($assignment.target.groupId -eq $Target.groupId) {
            if ($assignment.target.'@odata.type' -eq $Target.'@odata.type') {
                return @{ Status = "AlreadyExists"; Before = $before; After = $before }
            }
            return @{ Status = "Conflict"; Before = $before; After = $before }
        }
    }

    # Carry every existing target across verbatim. This preserves All Devices,
    # All Users, ConfigMgr collection targets, and any filters already in place.
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($assignment in $existing) { $targets.Add($assignment.target) }
    $targets.Add($Target)

    Set-ProfileAssignment -Kind $Kind -Id $Id -Targets $targets.ToArray()
    return @{ Status = "Added"; Before = $before; After = $targets.Count }
}

# Remove every target for a group from a configuration profile, preserving all
# other targets. Returns a hashtable: Status, Before, After, Removed.
function Remove-ProfileAssignment {
    param (
        [string]$Kind,
        [string]$Id,
        [string]$GroupId
    )

    # See the note in Add-ProfileAssignment - do not wrap this in @().
    $existing = Get-ProfileAssignment -Kind $Kind -Id $Id
    $before   = $existing.Count

    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($assignment in $existing) {
        if ($assignment.target.groupId -ne $GroupId) { $targets.Add($assignment.target) }
    }

    $removed = $before - $targets.Count
    if ($removed -eq 0) {
        return @{ Status = "NotAssigned"; Before = $before; After = $before; Removed = 0 }
    }

    Set-ProfileAssignment -Kind $Kind -Id $Id -Targets $targets.ToArray()
    return @{ Status = "Removed"; Before = $before; After = $targets.Count; Removed = $removed }
}

# ---------------------------------------------------------------------------
# Helper functions - scope tags (read-merge-write, additive only)
# ---------------------------------------------------------------------------

# Add or remove scope tags on an object, preserving every tag not named.
#
# roleScopeTagIds is written as a whole array, so both directions are a
# read-merge-write: read the current tags, apply the change, write the result.
# The tag list is never replaced with just the selection.
#
# Returns a hashtable: Status, Before, After.
#   Status "Updated"           - the tag list changed
#   Status "NoChange"          - the requested change was already the case
#   Status "SkippedWouldEmpty" - removal refused; it would have left no tags
function Update-ObjectScopeTag {
    param (
        [string]$Kind,
        [string]$Id,
        [string]$OdataType,
        [string[]]$TagIds,
        [ValidateSet("Add", "Remove")]
        [string]$Action = "Add"
    )

    if ($Kind -eq "mobileApps") {
        $uri = "$graphBeta/deviceAppManagement/mobileApps/$Id"
    } else {
        $uri = Get-ProfileUri -Kind $Kind -Id $Id
    }

    # Prefer the narrow projection, but fall back to a full entity read if the
    # endpoint rejects $select for this type.
    try {
        $current = Invoke-MgGraphRequest -Method GET -Uri "$uri`?`$select=id,roleScopeTagIds" -ErrorAction Stop
    } catch {
        Write-LogMessage "Narrow scope tag read failed for $Kind/$Id ($_); retrying without `$select." "WARN"
        $current = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    }

    $existing = @($current.roleScopeTagIds | ForEach-Object { "$_" })
    $before   = $existing.Count

    $merged = New-Object System.Collections.Generic.List[string]
    if ($Action -eq "Remove") {
        foreach ($tag in $existing) {
            if ($TagIds -notcontains $tag) { $merged.Add($tag) }
        }

        # An object with no scope tags is not visible to any scoped administrator,
        # and getting it back requires an unscoped account. Removing the last tag
        # is refused rather than performed - the operator can still remove it
        # deliberately in the portal, where the consequence is theirs to see.
        if ($merged.Count -eq 0 -and $before -gt 0) {
            return @{ Status = "SkippedWouldEmpty"; Before = $before; After = $before }
        }
    } else {
        foreach ($tag in $existing) { if ($merged -notcontains $tag) { $merged.Add($tag) } }
        foreach ($tag in $TagIds)   { if ($merged -notcontains $tag) { $merged.Add($tag) } }
    }

    if ($merged.Count -eq $before) {
        return @{ Status = "NoChange"; Before = $before; After = $before }
    }

    $body = @{ roleScopeTagIds = $merged.ToArray() }

    # These Graph entity types are polymorphic - Graph rejects a PATCH that does
    # not declare the concrete @odata.type. configurationPolicies and
    # groupPolicyConfigurations are not polymorphic and must not carry one.
    if (-not [string]::IsNullOrWhiteSpace($OdataType)) {
        $body."@odata.type" = $OdataType
    }

    $json = $body | ConvertTo-Json -Depth 10
    Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop | Out-Null
    return @{ Status = "Updated"; Before = $before; After = $merged.Count }
}

# ---------------------------------------------------------------------------
# Helper functions - UI
# ---------------------------------------------------------------------------

# Stop a sorted column's header from rendering as if it were selected.
#
# A DataGridView draws the header of the currently sorted column using the grid's
# SELECTION colours, so clicking a header to sort leaves it highlighted in the
# same blue as a selected row - which reads as "this column is selected" when
# nothing of the sort is true. Overriding the header selection colours to match
# the normal header colours removes the highlight and keeps sorting working.
# EnableHeadersVisualStyles must be off, or the theme paints the headers and
# these values are ignored.
function Set-GridHeaderStyle {
    param ([System.Windows.Forms.DataGridView]$Grid)

    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersDefaultCellStyle.BackColor          = [System.Drawing.SystemColors]::Control
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor          = [System.Drawing.SystemColors]::ControlText
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.SystemColors]::Control
    $Grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.SystemColors]::ControlText
}

# Populate the group grid.
function Set-GroupGridData {
    param (
        [System.Windows.Forms.DataGridView]$Grid,
        [object[]]$Items
    )
    $Grid.Rows.Clear()
    foreach ($item in $Items) {
        $rowIndex = $Grid.Rows.Add()
        $Grid.Rows[$rowIndex].Cells["DisplayName"].Value = $item.DisplayName
        $Grid.Rows[$rowIndex].Cells["Id"].Value          = $item.Id
    }
}

# Run a scriptblock with a wait cursor so long Graph calls do not look hung.
#
# IMPORTANT: the scriptblock is invoked with "&", which runs it in a CHILD scope.
# Assignments to a caller's local variables inside the block are lost. Use this
# only for blocks that return a value, mutate an existing object (for example
# List.Add), or write to $global: variables. For a block that updates local
# counters, set the cursor inline with a try/finally instead.
function Invoke-WithWaitCursor {
    param ([scriptblock]$Action)
    $previous = [System.Windows.Forms.Cursor]::Current
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        & $Action
    } finally {
        [System.Windows.Forms.Cursor]::Current = $previous
    }
}

# ---------------------------------------------------------------------------
# Form and tab scaffolding
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = "Intune Assignment Wizard"
$form.ClientSize    = New-Object System.Drawing.Size(764, 620)
$form.StartPosition = "CenterScreen"
$form.MinimumSize   = New-Object System.Drawing.Size(780, 660)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($tabControl)

# Tab indices: 0=Connect, 1=Objects, 2=Assignment, 3=Scope Tags, 4=Summary
$tabConnect = New-Object System.Windows.Forms.TabPage; $tabConnect.Text = "Connect"
$tabObjects = New-Object System.Windows.Forms.TabPage; $tabObjects.Text = "Objects"
$tabAssignment = New-Object System.Windows.Forms.TabPage; $tabAssignment.Text = "Assignment"
$tabScopeTags = New-Object System.Windows.Forms.TabPage; $tabScopeTags.Text = "Scope Tags"
$tabSummary = New-Object System.Windows.Forms.TabPage; $tabSummary.Text = "Summary"

# Every decision about WHAT to change - group, action, filter, scope tags - comes
# before the object list. For Remove, building that list means surveying the
# estate, which is the one long wait in the tool; asking the remaining questions
# after it would strand the operator mid-task. It also means the object list is
# the last input, so nothing invalidates it once it has been built.
$tabControl.TabPages.Add($tabConnect)
$tabControl.TabPages.Add($tabAssignment)
$tabControl.TabPages.Add($tabScopeTags)
$tabControl.TabPages.Add($tabObjects)
$tabControl.TabPages.Add($tabSummary)

$labelFont = New-Object System.Drawing.Font("Segoe UI", 10)
$monoFont  = New-Object System.Drawing.Font("Consolas", 10)

# Standard bottom button row
$buttonRowY = 552

# ---------------------------------------------------------------------------
# Tab 1 - Connect: tenant, client, mode, and authentication
# ---------------------------------------------------------------------------
$tenantIdLabel = New-Object System.Windows.Forms.Label
$tenantIdLabel.Text     = "Tenant ID:"
$tenantIdLabel.Location = New-Object System.Drawing.Point(10, 20)
$tenantIdLabel.Size     = New-Object System.Drawing.Size(120, 25)
$tenantIdLabel.Font     = $labelFont
$tabConnect.Controls.Add($tenantIdLabel)

$tenantIdTextBox = New-Object System.Windows.Forms.TextBox
$tenantIdTextBox.Location = New-Object System.Drawing.Point(140, 20)
$tenantIdTextBox.Size     = New-Object System.Drawing.Size(600, 25)
$tenantIdTextBox.Font     = $labelFont
$tenantIdTextBox.Text     = $defaultTenantId
$tabConnect.Controls.Add($tenantIdTextBox)

$clientIdLabel = New-Object System.Windows.Forms.Label
$clientIdLabel.Text     = "Client ID:"
$clientIdLabel.Location = New-Object System.Drawing.Point(10, 60)
$clientIdLabel.Size     = New-Object System.Drawing.Size(120, 25)
$clientIdLabel.Font     = $labelFont
$tabConnect.Controls.Add($clientIdLabel)

$clientIdTextBox = New-Object System.Windows.Forms.TextBox
$clientIdTextBox.Location = New-Object System.Drawing.Point(140, 60)
$clientIdTextBox.Size     = New-Object System.Drawing.Size(600, 25)
$clientIdTextBox.Font     = $labelFont
$clientIdTextBox.Text     = $defaultClientId
$tabConnect.Controls.Add($clientIdTextBox)

$authHintLabel = New-Object System.Windows.Forms.Label
$authHintLabel.Text      = "Both are optional. Leave Client ID blank to use the built-in Microsoft Graph PowerShell app; leave Tenant ID blank to use your home tenant."
$authHintLabel.Location  = New-Object System.Drawing.Point(140, 90)
$authHintLabel.Size      = New-Object System.Drawing.Size(610, 36)
$authHintLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$authHintLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
$tabConnect.Controls.Add($authHintLabel)

# Guards against this tool escalating an app registration it was merely pointed
# at. See the comment in Connect-GraphSession for why this matters.
$grantedOnlyCheckBox = New-Object System.Windows.Forms.CheckBox
$grantedOnlyCheckBox.Text     = "Do not request new permissions - use only what this app has already been granted"
$grantedOnlyCheckBox.Location = New-Object System.Drawing.Point(140, 132)
$grantedOnlyCheckBox.Size     = New-Object System.Drawing.Size(610, 24)
$grantedOnlyCheckBox.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grantedOnlyCheckBox.Checked  = (-not [string]::IsNullOrWhiteSpace($defaultClientId))
$grantedOnlyCheckBox.Enabled  = (-not [string]::IsNullOrWhiteSpace($defaultClientId))
$tabConnect.Controls.Add($grantedOnlyCheckBox)

# Only meaningful for a named app registration. The built-in Microsoft Graph
# PowerShell app is shared across every tool that uses it and holds no standing
# grant of its own, so it must request its scopes explicitly.
$clientIdTextBox.Add_TextChanged({
    $hasClientId = -not [string]::IsNullOrWhiteSpace($clientIdTextBox.Text)
    $grantedOnlyCheckBox.Enabled = $hasClientId
    if (-not $hasClientId) { $grantedOnlyCheckBox.Checked = $false }
    else { $grantedOnlyCheckBox.Checked = $true }
})

# Mode drives which object type the rest of the wizard operates on.
# Two independent axes: what kind of object, and what to change about it. Two
# dropdowns rather than one combined list - the axes do not constrain each other,
# so a single list would multiply out to twelve entries describing four choices.
# The consequences of each pairing are explained by the hint below, which
# rewrites itself as the selections change, so nothing is lost by collapsing the
# radio buttons that previously carried that text.
$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text     = "Work with:"
$modeLabel.Location = New-Object System.Drawing.Point(10, 178)
$modeLabel.Size     = New-Object System.Drawing.Size(120, 25)
$modeLabel.Font     = $labelFont
$tabConnect.Controls.Add($modeLabel)

$modeComboBox = New-Object System.Windows.Forms.ComboBox
$modeComboBox.Location      = New-Object System.Drawing.Point(140, 175)
$modeComboBox.Size          = New-Object System.Drawing.Size(280, 25)
$modeComboBox.Font          = $labelFont
$modeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$modeComboBox.Items.Add("Apps")
[void]$modeComboBox.Items.Add("Configuration profiles")
[void]$modeComboBox.Items.Add("Both apps and profiles")
$modeComboBox.SelectedIndex = 0
$tabConnect.Controls.Add($modeComboBox)

$changeLabel = New-Object System.Windows.Forms.Label
$changeLabel.Text     = "Change:"
$changeLabel.Location = New-Object System.Drawing.Point(440, 178)
$changeLabel.Size     = New-Object System.Drawing.Size(90, 25)
$changeLabel.Font     = $labelFont
$tabConnect.Controls.Add($changeLabel)

$changeComboBox = New-Object System.Windows.Forms.ComboBox
$changeComboBox.Location      = New-Object System.Drawing.Point(530, 175)
$changeComboBox.Size          = New-Object System.Drawing.Size(220, 25)
$changeComboBox.Font          = $labelFont
$changeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$changeComboBox.Items.Add("Assignments only")
[void]$changeComboBox.Items.Add("Scope tags only")
[void]$changeComboBox.Items.Add("Assignments and scope tags")
$changeComboBox.SelectedIndex = 0
$tabConnect.Controls.Add($changeComboBox)

$selectionHintLabel = New-Object System.Windows.Forms.Label
$selectionHintLabel.Location  = New-Object System.Drawing.Point(10, 210)
$selectionHintLabel.Size      = New-Object System.Drawing.Size(740, 56)
$selectionHintLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$selectionHintLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
$tabConnect.Controls.Add($selectionHintLabel)

# Explain the pairing that is actually selected, rather than every pairing at
# once. The cost of a removal is the part worth knowing before committing to it.
function Update-SelectionHint {
    switch ($modeComboBox.SelectedIndex) {
        1       { $modeText = "Configuration profiles, including compliance policies and remediation scripts, for iOS, Android, macOS, and Windows." }
        2       { $modeText = "Apps and configuration profiles, including compliance policies and remediation scripts, for iOS, Android, macOS, and Windows." }
        default { $modeText = "Apps for iOS, Android, macOS, and Windows." }
    }

    switch ($changeComboBox.SelectedIndex) {
        1       { $changeText = "No group is involved - the Assignment tab is skipped. Removing tags finds objects by the tags they already carry, which needs no survey." }
        2       { $changeText = "One group change and one scope tag change, applied to the same objects in a single pass." }
        default { $changeText = "The Scope Tags tab is skipped. Removing a group has to survey every object in scope, which takes a while on a large tenant." }
    }

    $selectionHintLabel.Text = "$modeText  $changeText"

    # Any object list built under different selections is no longer valid, so
    # drop it rather than let the Objects tab reuse it. The context key would
    # catch this too, but clearing here means the change takes effect the moment
    # it is made rather than at the next visit.
    $global:DiscoveryContext = ""
}

# Read both selectors into the globals the rest of the wizard uses.
#
# These were previously read only inside the Authenticate handler, so changing a
# dropdown afterwards had no effect: the run kept whatever had been selected at
# the last sign-in. Resolving them on every change - and again before the
# selections are used - keeps the globals and the controls in step.
function Update-SelectionState {
    switch ($modeComboBox.SelectedIndex) {
        1       { $global:Mode = "Profiles" }
        2       { $global:Mode = "Both" }
        default { $global:Mode = "Apps" }
    }

    switch ($changeComboBox.SelectedIndex) {
        1       { $global:ChangeScope = "ScopeTags" }
        2       { $global:ChangeScope = "Both" }
        default { $global:ChangeScope = "Assignments" }
    }

    Update-SelectionHint
}

$modeComboBox.Add_SelectedIndexChanged({ Update-SelectionState })
$changeComboBox.Add_SelectedIndexChanged({ Update-SelectionState })
Update-SelectionHint

$connectInfoLabel = New-Object System.Windows.Forms.Label
$connectInfoLabel.Text     = "This wizard never replaces an object's assignments or scope tags. It applies only the change you ask for and leaves everything else on the object untouched."
$connectInfoLabel.Location = New-Object System.Drawing.Point(10, 276)
$connectInfoLabel.Size     = New-Object System.Drawing.Size(730, 44)
$connectInfoLabel.Font     = $labelFont
$tabConnect.Controls.Add($connectInfoLabel)

$logPathLabel = New-Object System.Windows.Forms.Label
$logPathLabel.Text     = "Log file: $global:LogPath"
$logPathLabel.Location = New-Object System.Drawing.Point(10, 328)
$logPathLabel.Size     = New-Object System.Drawing.Size(730, 25)
$tabConnect.Controls.Add($logPathLabel)

$authStatusLabel = New-Object System.Windows.Forms.Label
$authStatusLabel.Text     = "Not connected."
$authStatusLabel.Location = New-Object System.Drawing.Point(10, 358)
$authStatusLabel.Size     = New-Object System.Drawing.Size(730, 50)
$authStatusLabel.Font     = $labelFont
$tabConnect.Controls.Add($authStatusLabel)

$authenticateButton = New-Object System.Windows.Forms.Button
$authenticateButton.Location = New-Object System.Drawing.Point(10, $buttonRowY)
$authenticateButton.Size     = New-Object System.Drawing.Size(140, 35)
$authenticateButton.Text     = "Authenticate"
$authenticateButton.Font     = $labelFont
$authenticateButton.Add_Click({
    $global:ClientId = $clientIdTextBox.Text.Trim()
    $global:TenantId = $tenantIdTextBox.Text.Trim()
    Update-SelectionState


    Write-LogMessage "Checking Graph authentication module..."
    if (-not (Assert-GraphAuthModule)) {
        [System.Windows.Forms.MessageBox]::Show("Could not load Microsoft.Graph.Authentication. See log: $global:LogPath", "Module Error")
        return
    }

    $global:LastAuthError = $null
    $global:LastAuthHint  = $null
    Write-LogMessage "Authenticating to Microsoft Graph..."
    $global:UseGrantedPermissionsOnly = $grantedOnlyCheckBox.Checked
    $context = Invoke-WithWaitCursor {
        Connect-GraphSession -ClientId $global:ClientId -TenantId $global:TenantId `
                             -UseGrantedPermissionsOnly:$global:UseGrantedPermissionsOnly
    }
    if (-not $context) {
        $authMessage = "Authentication failed."
        if ($global:LastAuthHint) {
            $authMessage += [Environment]::NewLine + [Environment]::NewLine + $global:LastAuthHint
        }
        $authMessage += [Environment]::NewLine + [Environment]::NewLine + "Full detail in the log: $global:LogPath"
        [System.Windows.Forms.MessageBox]::Show($authMessage, "Authentication Error")
        return
    }

    $authStatusLabel.Text      = "Connected as $(Get-GraphIdentityLabel $context)   |   tenant $($context.TenantId)"
    $authStatusLabel.ForeColor = [System.Drawing.SystemColors]::ControlText

    # A successful sign-in does NOT mean the token carries the scopes this tool
    # needs. Reads commonly succeed while writes return 403 much later, after the
    # operator has already worked through the wizard. Check up front instead.
    $problems = New-Object System.Collections.Generic.List[string]

    # Check 1 - delegated scopes on the token. These come from the app
    # registration's consent.
    $grantedScopes = @($context.Scopes)
    $missingScopes = @($global:RequiredScopes | Where-Object { $grantedScopes -notcontains $_ })
    if ($missingScopes.Count -gt 0) {
        Write-LogMessage "Token is missing required scope(s): $($missingScopes -join ', '). Writes will fail with 403 until these are consented for the app registration in use." "WARN"
        Write-LogMessage "Scopes actually granted: $($grantedScopes -join ', ')"
        $problems.Add("Missing delegated Graph scopes (granted by app registration consent):")
        foreach ($scope in $missingScopes) { $problems.Add("    $scope") }
        if ($global:UseGrantedPermissionsOnly) {
            $problems.Add("")
            $problems.Add("This session requested only the permissions the app already holds,")
            $problems.Add("so nothing was added to it. Grant these scopes to the app")
            $problems.Add("registration in Entra, or clear the checkbox to be prompted to")
            $problems.Add("consent to them - which permanently adds them to this app.")
        }
    } else {
        Write-LogMessage "All required delegated scopes are present on the token."
    }

    # Always record what the token actually carries, not only when something is
    # missing. When diagnosing a 403 after the fact, the granted scope list is
    # the first thing worth seeing, and it also makes clear which scopes were
    # obtained by dynamic consent rather than pre-listed on the registration.
    Write-LogMessage "Token scopes granted: $($grantedScopes -join ', ')"

    # Check 2 - Intune RBAC. Independent of the scopes above: this is what the
    # account's Intune role actually permits. Advisory only, because the
    # endpoint is not available in every tenant and the action names have
    # changed over time - an inconclusive result must not raise a false alarm.
    $allowedActions = Invoke-WithWaitCursor { Get-IntuneEffectivePermission }
    $global:AllowedIntuneActions = $allowedActions
    if ($allowedActions) {
        # Both mode spans two resource families, each with its own action names.
        if ($global:Mode -eq "Both") {
            $modesToCheck = @("Apps", "Profiles")
        } elseif ($global:Mode -eq "Profiles") {
            $modesToCheck = @("Profiles", "AppConfigurations")
        } else {
            $modesToCheck = @($global:Mode)
        }

        $missingActions = New-Object System.Collections.Generic.List[string]
        $anyConclusive  = $false
        foreach ($modeToCheck in $modesToCheck) {
            $permissionCheck = Get-MissingIntunePermission -AllowedActions $allowedActions -Mode $modeToCheck
            if ($permissionCheck.Conclusive) {
                $anyConclusive = $true
                foreach ($action in @($permissionCheck.Missing)) { $missingActions.Add($action) }
            }
        }

        if (-not $anyConclusive) {
            Write-LogMessage "Intune RBAC check was inconclusive for mode '$global:Mode'; no permissions warning raised."
        } elseif ($missingActions.Count -gt 0) {
            Write-LogMessage "Account's Intune role does not appear to grant: $($missingActions -join '; ')" "WARN"
            $problems.Add("Intune RBAC role does not appear to allow:")
            foreach ($action in $missingActions) { $problems.Add("    $action") }
        } else {
            Write-LogMessage "Account's Intune role grants the resource actions needed for mode '$global:Mode'."
        }
    }

    if ($problems.Count -gt 0) {
        $authStatusLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $authStatusLabel.Text = "Connected as $(Get-GraphIdentityLabel $context)   |   INSUFFICIENT PERMISSIONS - writes will fail"

        $warning = "Signed in successfully, but this session is missing permissions needed to make changes." +
                   [Environment]::NewLine + [Environment]::NewLine +
                   ($problems -join [Environment]::NewLine) +
                   [Environment]::NewLine + [Environment]::NewLine +
                   "Reads will work. Assignment and scope tag changes will fail with 403 Forbidden."
        if ([string]::IsNullOrWhiteSpace($global:ClientId)) {
            $warning += [Environment]::NewLine + [Environment]::NewLine +
                        "No Client ID was supplied, so the built-in Microsoft Graph PowerShell app is in use. Many tenants allow that app to read Intune but not to write to it. Enter your own app registration's Client ID on the Connect tab and authenticate again."
        }
        $warning += [Environment]::NewLine + [Environment]::NewLine + "You can continue, but writes are expected to fail."
        [System.Windows.Forms.MessageBox]::Show($warning, "Insufficient Permissions")
    }
    Write-LogMessage "Successfully authenticated. Mode: $global:Mode."

    # Adapt the Objects tab to the chosen mode.
    if ($global:Mode -eq "Apps") {
        $objectsLabel.Text = "Search apps by display-name prefix, then select one or more:"
    } else {
        $objectsLabel.Text = "Search app and device configuration profiles, compliance policies, and remediation scripts by display-name prefix, then select one or more:"
    }
    $objectGridView.Rows.Clear()
    $objectCountLabel.Text = ""
    $selectAllButton.Text  = "Select All"

    # Load the reference data used by the Assignment and Scope Tags tabs.
    try {
        Invoke-WithWaitCursor {
            $global:AssignmentFilters = @(Get-WindowsAssignmentFilters)
            $global:ScopeTags         = @(Get-RoleScopeTags)
        }
        $filterComboBox.Items.Clear()
        [void]$filterComboBox.Items.Add("(None)")
        foreach ($assignmentFilter in $global:AssignmentFilters) {
            [void]$filterComboBox.Items.Add($assignmentFilter.DisplayName)
        }
        $filterComboBox.SelectedIndex = 0

        $scopeTagCheckedListBox.Items.Clear()
        foreach ($scopeTag in $global:ScopeTags) {
            [void]$scopeTagCheckedListBox.Items.Add("$($scopeTag.DisplayName)  (ID $($scopeTag.Id))")
        }
        Write-LogMessage "Loaded $($global:AssignmentFilters.Count) Windows assignment filter(s) and $($global:ScopeTags.Count) scope tag(s)."
    } catch {
        Write-LogMessage "Error loading filters or scope tags: $_" "WARN"
        [System.Windows.Forms.MessageBox]::Show("Connected, but filters or scope tags could not be loaded. Check the log and confirm the DeviceManagementRBAC.Read.All scope was consented.", "Partial Load")
    }

    # Authentication no longer jumps to the next tab. The permission warnings
    # raised above are worth reading before moving on, and re-authenticating
    # against a different app registration or tenant should not shunt the
    # operator forward each time.
    $nextButton1.Enabled = $true
})
$tabConnect.Controls.Add($authenticateButton)

$nextButton1 = New-Object System.Windows.Forms.Button
$nextButton1.Location = New-Object System.Drawing.Point(640, $buttonRowY)
$nextButton1.Size     = New-Object System.Drawing.Size(110, 35)
$nextButton1.Text     = "Next"
$nextButton1.Font     = $labelFont
$nextButton1.Enabled  = $false
$nextButton1.Add_Click({
    if (-not (Get-MgContext)) {
        [System.Windows.Forms.MessageBox]::Show("Authenticate before continuing.", "Not Connected")
        return
    }
    # The dropdowns may have changed since the last sign-in, so resolve them
    # again here and say what the run will actually do. Without this line the
    # only record of the mode is the one written at authentication, which makes
    # a later change impossible to confirm from the log.
    $previousMode = $global:Mode
    Update-SelectionState
    Write-LogMessage "Proceeding with mode '$($global:Mode)', changing '$($global:ChangeScope)'."

    # The Intune RBAC check ran against the mode selected at sign-in. If the mode
    # has changed it covers a different resource family, so repeat it - the
    # effective permissions were kept, so this costs nothing.
    if ($previousMode -ne $global:Mode -and $global:AllowedIntuneActions) {
        if ($global:Mode -eq "Both") {
            $modesToCheck = @("Apps", "Profiles")
        } elseif ($global:Mode -eq "Profiles") {
            $modesToCheck = @("Profiles", "AppConfigurations")
        } else {
            $modesToCheck = @($global:Mode)
        }
        $newMissing = New-Object System.Collections.Generic.List[string]
        foreach ($modeToCheck in $modesToCheck) {
            $check = Get-MissingIntunePermission -AllowedActions $global:AllowedIntuneActions -Mode $modeToCheck
            if ($check.Conclusive) {
                foreach ($action in @($check.Missing)) { $newMissing.Add($action) }
            }
        }
        if ($newMissing.Count -gt 0) {
            Write-LogMessage "Mode changed to '$($global:Mode)'; the account's Intune role does not appear to grant: $($newMissing -join '; ')" "WARN"
            [System.Windows.Forms.MessageBox]::Show(
                "The mode changed since sign-in, and the Intune role does not appear to allow:" + [Environment]::NewLine + [Environment]::NewLine +
                "    " + ($newMissing -join ([Environment]::NewLine + "    ")) + [Environment]::NewLine + [Environment]::NewLine +
                "Reads will work; writes are expected to fail with 403.",
                "Insufficient Permissions")
        } else {
            Write-LogMessage "Mode changed to '$($global:Mode)'; the account's Intune role covers it."
        }
    }

    Update-TabVisibility

    if ($global:ChangeScope -eq "ScopeTags") {
        # No group is involved, so the Assignment tab is skipped entirely and the
        # direction comes from the selector on the Scope Tags tab.
        if ($scopeTagRemoveRadio.Checked) { $global:Action = "Remove" } else { $global:Action = "Add" }
        $global:GroupId       = $null
        $global:SelectedGroup = $null
        Initialize-ScopeTagsTab
        Set-WizardTab -Target $tabScopeTags -Direction "Next"
    } else {
        Update-RemoveScaleNote
        Set-WizardTab -Target $tabAssignment -Direction "Next"
    }
})
$tabConnect.Controls.Add($nextButton1)

$cancelButton1 = New-Object System.Windows.Forms.Button
$cancelButton1.Location = New-Object System.Drawing.Point(530, $buttonRowY)
$cancelButton1.Size     = New-Object System.Drawing.Size(100, 35)
$cancelButton1.Text     = "Cancel"
$cancelButton1.Font     = $labelFont
$cancelButton1.Add_Click({ $form.Close() })
$tabConnect.Controls.Add($cancelButton1)

# ---------------------------------------------------------------------------
# Tab 2 - Objects: search and multi-select apps or profiles
# ---------------------------------------------------------------------------
$objectsLabel = New-Object System.Windows.Forms.Label
$objectsLabel.Text     = "Load objects by name prefix, narrow them with the filter row, then select one or more:"
$objectsLabel.Location = New-Object System.Drawing.Point(10, 10)
$objectsLabel.Size     = New-Object System.Drawing.Size(740, 25)
$objectsLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectsLabel)

$objectPrefixLabel = New-Object System.Windows.Forms.Label
$objectPrefixLabel.Text     = "Name starts with:"
$objectPrefixLabel.Location = New-Object System.Drawing.Point(10, 45)
$objectPrefixLabel.Size     = New-Object System.Drawing.Size(130, 25)
$objectPrefixLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectPrefixLabel)

$objectPrefixTextBox = New-Object System.Windows.Forms.TextBox
$objectPrefixTextBox.Location = New-Object System.Drawing.Point(145, 43)
$objectPrefixTextBox.Size     = New-Object System.Drawing.Size(360, 25)
$objectPrefixTextBox.Font     = $labelFont
$objectPrefixTextBox.Text     = $defaultObjectPrefix
$tabObjects.Controls.Add($objectPrefixTextBox)

$loadObjectsButton = New-Object System.Windows.Forms.Button
$loadObjectsButton.Location = New-Object System.Drawing.Point(515, 41)
$loadObjectsButton.Size     = New-Object System.Drawing.Size(90, 30)
$loadObjectsButton.Text     = "Load"
$loadObjectsButton.Font     = $labelFont
$tabObjects.Controls.Add($loadObjectsButton)

$objectPrefixHint = New-Object System.Windows.Forms.Label
$objectPrefixHint.Text     = "Blank = all"
$objectPrefixHint.Location = New-Object System.Drawing.Point(615, 48)
$objectPrefixHint.Size     = New-Object System.Drawing.Size(130, 25)
$tabObjects.Controls.Add($objectPrefixHint)

# Second row: client-side narrowing of the already-loaded results. This never
# calls Graph - it only filters what the Load button retrieved, so it is instant
# and can be changed freely without re-querying.
$objectFilterLabel = New-Object System.Windows.Forms.Label
$objectFilterLabel.Text     = "Filter list:"
$objectFilterLabel.Location = New-Object System.Drawing.Point(10, 82)
$objectFilterLabel.Size     = New-Object System.Drawing.Size(70, 25)
$objectFilterLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectFilterLabel)

$objectFilterTextBox = New-Object System.Windows.Forms.TextBox
$objectFilterTextBox.Location = New-Object System.Drawing.Point(80, 80)
$objectFilterTextBox.Size     = New-Object System.Drawing.Size(230, 25)
$objectFilterTextBox.Font     = $labelFont
$tabObjects.Controls.Add($objectFilterTextBox)

$objectTypeLabel = New-Object System.Windows.Forms.Label
$objectTypeLabel.Text     = "Type:"
$objectTypeLabel.Location = New-Object System.Drawing.Point(320, 82)
$objectTypeLabel.Size     = New-Object System.Drawing.Size(40, 25)
$objectTypeLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectTypeLabel)

$objectTypeComboBox = New-Object System.Windows.Forms.ComboBox
$objectTypeComboBox.Location      = New-Object System.Drawing.Point(360, 79)
$objectTypeComboBox.Size          = New-Object System.Drawing.Size(140, 25)
$objectTypeComboBox.Font          = $labelFont
$objectTypeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabObjects.Controls.Add($objectTypeComboBox)

$objectPlatformLabel = New-Object System.Windows.Forms.Label
$objectPlatformLabel.Text     = "Platform:"
$objectPlatformLabel.Location = New-Object System.Drawing.Point(510, 82)
$objectPlatformLabel.Size     = New-Object System.Drawing.Size(60, 25)
$objectPlatformLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectPlatformLabel)

$objectPlatformComboBox = New-Object System.Windows.Forms.ComboBox
$objectPlatformComboBox.Location      = New-Object System.Drawing.Point(575, 79)
$objectPlatformComboBox.Size          = New-Object System.Drawing.Size(175, 25)
$objectPlatformComboBox.Font          = $labelFont
$objectPlatformComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$objectPlatformComboBox.Items.Add("(All platforms)")
[void]$objectPlatformComboBox.Items.Add("iOS")
[void]$objectPlatformComboBox.Items.Add("Android")
[void]$objectPlatformComboBox.Items.Add("macOS")
[void]$objectPlatformComboBox.Items.Add("Windows")
$objectPlatformComboBox.SelectedIndex = 0
$tabObjects.Controls.Add($objectPlatformComboBox)

$objectGridView = New-Object System.Windows.Forms.DataGridView
$objectGridView.Location           = New-Object System.Drawing.Point(10, 115)
$objectGridView.Size               = New-Object System.Drawing.Size(740, 425)
$objectGridView.AllowUserToAddRows = $false
$objectGridView.RowHeadersVisible  = $false
$objectGridView.SelectionMode      = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

$colObjSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colObjSelect.Name       = "Select"
$colObjSelect.HeaderText = "Select"
$colObjSelect.Width      = 55
$objectGridView.Columns.Add($colObjSelect) | Out-Null

$colObjName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjName.Name       = "DisplayName"
$colObjName.HeaderText = "Display Name"
$colObjName.Width      = 250
$colObjName.ReadOnly   = $true
$objectGridView.Columns.Add($colObjName) | Out-Null

$colObjTypeLabel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjTypeLabel.Name       = "TypeLabel"
$colObjTypeLabel.HeaderText = "Type"
$colObjTypeLabel.Width      = 120
$colObjTypeLabel.ReadOnly   = $true
$objectGridView.Columns.Add($colObjTypeLabel) | Out-Null

$colObjPlatform = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjPlatform.Name       = "Platform"
$colObjPlatform.HeaderText = "Platform"
$colObjPlatform.Width      = 85
$colObjPlatform.ReadOnly   = $true
$objectGridView.Columns.Add($colObjPlatform) | Out-Null

# Populated only in Remove mode, where how the group is currently targeted is
# the decisive detail - an exclusion or a sole assignment changes what removal
# actually does.
$colObjAssignment = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjAssignment.Name       = "AssignmentSummary"
$colObjAssignment.HeaderText = "Current Assignment"
$colObjAssignment.Width      = 230
$colObjAssignment.ReadOnly   = $true
$objectGridView.Columns.Add($colObjAssignment) | Out-Null

$colObjId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjId.Name       = "Id"
$colObjId.HeaderText = "Id"
$colObjId.Width      = 155
$colObjId.ReadOnly   = $true
$colObjId.Visible    = $false
$objectGridView.Columns.Add($colObjId) | Out-Null

# Hidden columns carry the Graph entity set and @odata.type through to
# execution time so the correct endpoint and PATCH body can be built.
$colObjKind = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjKind.Name    = "Kind"
$colObjKind.Visible = $false
$objectGridView.Columns.Add($colObjKind) | Out-Null

$colObjOdataType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjOdataType.Name    = "OdataType"
$colObjOdataType.Visible = $false
$objectGridView.Columns.Add($colObjOdataType) | Out-Null
Set-GridHeaderStyle -Grid $objectGridView
$tabObjects.Controls.Add($objectGridView)

# Return the loaded objects matching the current free-text and type filters.
# Substring matching uses IndexOf with OrdinalIgnoreCase rather than -like so
# that regex/wildcard characters in a profile name (brackets, asterisks) are
# treated literally instead of altering the match.
function Get-FilteredObjects {
    $text = $objectFilterTextBox.Text
    $type = [string]$objectTypeComboBox.SelectedItem
    $platform = [string]$objectPlatformComboBox.SelectedItem

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $global:Objects) {
        if ($type -and $type -ne "(All types)" -and $item.TypeLabel -ne $type) { continue }
        if ($platform -and $platform -ne "(All platforms)" -and $item.Platform -ne $platform) { continue }
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $name = "$($item.DisplayName)"
            if ($name.IndexOf($text, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }
        $results.Add($item)
    }
    return ,$results.ToArray()
}

# Update the "showing x of y, n selected" readout.
function Update-ObjectCountLabel {
    $shown = $objectGridView.Rows.Count
    $total = $global:Objects.Count
    $objectCountLabel.Text = "Showing $shown of $total   |   $($global:CheckedObjectIds.Count) selected"
}

# Re-render the grid from the current filters, restoring each row's checkbox
# from the selection set so ticks survive filter changes.
function Update-ObjectGrid {
    $global:SuppressGridEvents = $true
    try {
        $objectGridView.Rows.Clear()
        foreach ($item in (Get-FilteredObjects)) {
            $rowIndex = $objectGridView.Rows.Add()
            $row      = $objectGridView.Rows[$rowIndex]
            $row.Cells["Select"].Value      = $global:CheckedObjectIds.Contains([string]$item.Id)
            $row.Cells["DisplayName"].Value = $item.DisplayName
            $row.Cells["TypeLabel"].Value   = $item.TypeLabel
            $row.Cells["Platform"].Value   = $item.Platform
            $row.Cells["Id"].Value          = $item.Id
            $row.Cells["Kind"].Value        = $item.Kind
            $row.Cells["OdataType"].Value   = $item.OdataType

            if ($item.PSObject.Properties.Name -contains "AssignmentSummary") {
                $row.Cells["AssignmentSummary"].Value = $item.AssignmentSummary

                # Removing an exclusion widens scope; removing a sole assignment
                # orphans the object. Both are easy to do by accident in a bulk
                # operation, so they are marked rather than left to be read.
                if ($item.AssignmentSummary -like "*ONLY ASSIGNMENT*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                } elseif ($item.AssignmentSummary -like "Exclude*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
                }
            }
        }
    } finally {
        $global:SuppressGridEvents = $false
    }
    Update-ObjectCountLabel
}

# A DataGridViewCheckBoxCell does not raise CellValueChanged until the cell
# loses focus. Committing the edit as soon as it goes dirty makes the click
# register immediately, which matters because the row may be filtered away
# before focus ever moves.
$objectGridView.Add_CurrentCellDirtyStateChanged({
    if ($objectGridView.IsCurrentCellDirty -and
        $objectGridView.CurrentCell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) {
        $objectGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$objectGridView.Add_CellValueChanged({
    param($eventSender, $gridEvent)
    if ($global:SuppressGridEvents) { return }
    if ($gridEvent.RowIndex -lt 0 -or $gridEvent.ColumnIndex -lt 0) { return }
    if ($objectGridView.Columns[$gridEvent.ColumnIndex].Name -ne "Select") { return }

    $row = $objectGridView.Rows[$gridEvent.RowIndex]
    $id  = [string]$row.Cells["Id"].Value
    if ($row.Cells["Select"].Value -eq $true) {
        [void]$global:CheckedObjectIds.Add($id)
    } else {
        [void]$global:CheckedObjectIds.Remove($id)
    }
    Update-ObjectCountLabel
})

$objectFilterTextBox.Add_TextChanged({ Update-ObjectGrid })
$objectTypeComboBox.Add_SelectedIndexChanged({ Update-ObjectGrid })
$objectPlatformComboBox.Add_SelectedIndexChanged({ Update-ObjectGrid })

# The Scope Tags tab follows the action chosen on the Assignment tab: selected
# tags are added under Add and removed under Remove. Having one tab quietly do
# the opposite of the selected action would be the worst of the options.
function Initialize-ScopeTagsTab {
    $scopeTagActionGroupBox.Visible = ($global:ChangeScope -eq "ScopeTags")

    if ($global:Action -eq "Remove") {
        $scopeTagLabel.Text      = "Optionally select scope tags to REMOVE from every selected object:"
        $scopeTagLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $scopeTagHintLabel.Text  = "Tags not selected here are never touched, and nothing changes if no tags are selected. An object's LAST remaining tag is never removed - that object is skipped and reported, because an object with no scope tags is invisible to every scoped administrator. The built-in Default tag is ID 0."
    } else {
        $scopeTagLabel.Text      = "Optionally select scope tags to ADD to every selected object:"
        $scopeTagLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $scopeTagHintLabel.Text  = "Tags not selected here are never touched, and nothing changes if no tags are selected. The built-in Default tag is ID 0."
    }
}

# Called when the Objects tab is entered. In Add mode nothing changes. In Remove
# mode the prefix search is meaningless - the object set is whatever targets the
# group - so discovery runs automatically and the controls are relabelled.
function Initialize-ObjectsTab {
    $isRemove = ($global:Action -eq "Remove")

    $objectPrefixLabel.Visible   = -not $isRemove
    $objectPrefixTextBox.Visible = -not $isRemove
    $objectPrefixHint.Visible    = -not $isRemove
    $objectGridView.Columns["AssignmentSummary"].Visible = $isRemove

    # Going Next and then Back must not cost the operator their work. Rebuilding
    # the list unconditionally would re-run the whole survey - minutes on a large
    # tenant - and discard every deselection made on the way through. So the list
    # is rebuilt only when the inputs it was derived from have actually changed;
    # otherwise the previous list and selection are shown again as they were.
    $context = "$($global:Mode)|$($global:Action)|$($global:ChangeScope)|$($global:GroupId)|$((@($global:SelectedScopeTags | ForEach-Object { $_.Id }) | Sort-Object) -join ',')"
    if ($global:DiscoveryContext -eq $context -and $global:Objects.Count -gt 0) {
        Write-LogMessage "Objects tab re-entered with unchanged group, action, and mode; keeping the existing list and selection."
        Set-ObjectsTabLabels -IsRemove $isRemove
        Update-ObjectGrid
        return
    }

    # Inputs changed, so anything chosen against the old ones is stale.
    if ($global:Objects.Count -gt 0) {
        Write-LogMessage "Group, action, or mode changed since the object list was built; discarding the previous list and selection."
    }
    $global:Objects = @()
    $global:CheckedObjectIds.Clear()

    Set-ObjectsTabLabels -IsRemove $isRemove

    if ($isRemove) {
        Invoke-ObjectDiscovery
    } else {
        # Add mode does not auto-load; clear the grid and wait for Load.
        $global:SuppressGridEvents = $true
        try {
            $objectFilterTextBox.Text = ""
            $objectTypeComboBox.Items.Clear()
            [void]$objectTypeComboBox.Items.Add("(All types)")
            $objectTypeComboBox.SelectedIndex = 0
            $objectPlatformComboBox.SelectedIndex = 0
        } finally {
            $global:SuppressGridEvents = $false
        }
        $selectAllButton.Text = "Select All"
        Update-ObjectGrid
        $global:DiscoveryContext = $context
    }
}

# Label and position the Objects tab controls for the action in effect.
function Set-ObjectsTabLabels {
    param ([bool]$IsRemove)

    if ($IsRemove -and $global:ChangeScope -eq "ScopeTags") {
        $objectsLabel.Text          = "Objects carrying the selected scope tag(s). Clear any you want to keep:"
        $loadObjectsButton.Text     = "Refresh"
        $loadObjectsButton.Location = New-Object System.Drawing.Point(10, 41)
    } elseif ($IsRemove) {
        $objectsLabel.Text          = "Objects currently assigned to '$($global:SelectedGroup)'. Clear any you want to keep:"
        $loadObjectsButton.Text     = "Refresh"
        $loadObjectsButton.Location = New-Object System.Drawing.Point(10, 41)
    } else {
        if ($global:Mode -eq "Apps") {
            $objectsLabel.Text = "Search apps by display-name prefix, then select one or more:"
        } elseif ($global:Mode -eq "Profiles") {
            $objectsLabel.Text = "Search configuration profiles, compliance policies, and remediation scripts by display-name prefix, then select one or more:"
        } else {
            $objectsLabel.Text = "Search apps, configuration profiles, compliance policies, and remediation scripts by display-name prefix, then select one or more:"
        }
        $loadObjectsButton.Text     = "Load"
        $loadObjectsButton.Location = New-Object System.Drawing.Point(515, 41)
    }
}

# Populate the grid, by discovery in Remove mode or by prefix search in Add mode.
function Invoke-ObjectDiscovery {
    $isRemove    = ($global:Action -eq "Remove")
    $byScopeTag  = ($global:ChangeScope -eq "ScopeTags")

    try {
        if ($isRemove -and $byScopeTag) {
            # Tag membership arrives with the list call, so there is nothing to
            # survey and nothing to warn about.
            $discovery = Invoke-WithWaitCursor {
                Find-ObjectsWithScopeTag -TagIds @($global:SelectedScopeTags | ForEach-Object { $_.Id }) -Mode $global:Mode
            }
            $global:Objects = @($discovery.Matches)
        } elseif ($isRemove) {
            # List first. Reading assignments is the expensive half, and the
            # operator should be able to back out once the size is known.
            $allObjects = @(Invoke-WithWaitCursor { Get-AllSupportedObjects -Mode $global:Mode })

            if ($allObjects.Count -gt $global:DiscoveryPromptThreshold) {
                $batchCount = [Math]::Ceiling($allObjects.Count / 20)
                $promptText = "$($allObjects.Count) object(s) must be surveyed to find which ones target '$($global:SelectedGroup)'." + [Environment]::NewLine + [Environment]::NewLine +
                              "Graph offers no way to ask which objects target a group, so each object's assignments have to be read. Batching reduces that to about $batchCount request(s), but this may still take a while and the window will be unresponsive until it finishes."
                if ($global:Mode -eq "Both") {
                    $promptText += [Environment]::NewLine + [Environment]::NewLine +
                                   "Narrowing the mode on the Connect tab to just apps or just profiles roughly halves the work."
                }
                $promptText += [Environment]::NewLine + [Environment]::NewLine + "Continue?"

                $proceed = [System.Windows.Forms.MessageBox]::Show(
                    $promptText, "Large Survey", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($proceed -ne [System.Windows.Forms.DialogResult]::Yes) {
                    Write-LogMessage "Operator cancelled a survey of $($allObjects.Count) object(s)."
                    return
                }
            }

            $discovery = Invoke-WithWaitCursor {
                Find-ObjectsAssignedToGroup -GroupId $global:GroupId -Objects $allObjects
            }
            $global:Objects = @($discovery.Matches)

            if (@($discovery.Unreadable).Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "$(@($discovery.Unreadable).Count) of $($discovery.Surveyed) object(s) could not be read and were skipped." + [Environment]::NewLine + [Environment]::NewLine +
                    "They may or may not target this group. Review the log before treating this list as complete.",
                    "Incomplete Survey")
            }
        } else {
            $prefix = $objectPrefixTextBox.Text
            $loaded = Invoke-WithWaitCursor {
                $items = New-Object System.Collections.Generic.List[object]
                if ($global:Mode -eq "Apps" -or $global:Mode -eq "Both") {
                    Write-LogMessage "Retrieving supported mobile apps (prefix '$prefix')..."
                    foreach ($item in (Get-MobileApps -Prefix $prefix)) { $items.Add($item) }
                }
                if ($global:Mode -eq "Profiles" -or $global:Mode -eq "Both") {
                    Write-LogMessage "Retrieving supported configuration profiles and remediation scripts (prefix '$prefix')..."
                    foreach ($item in (Get-ConfigurationProfiles -Prefix $prefix)) { $items.Add($item) }
                }
                return ,$items.ToArray()
            }
            $global:Objects = @($loaded)
        }
    } catch {
        Write-LogMessage "Error retrieving objects: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error retrieving objects. See log: $global:LogPath", "Graph Error")
        return
    }

    $global:CheckedObjectIds.Clear()

    # In Remove mode everything found is a hit by definition, so pre-check it -
    # the operator's job is to deselect exceptions rather than rebuild the list.
    if ($isRemove -and ($global:Objects.Count -gt 0)) {
        foreach ($item in $global:Objects) { [void]$global:CheckedObjectIds.Add([string]$item.Id) }
        $selectAllButton.Text = "Deselect All"
    } else {
        $selectAllButton.Text = "Select All"
    }

    $global:SuppressGridEvents = $true
    try {
        $objectFilterTextBox.Text = ""
        $objectTypeComboBox.Items.Clear()
        [void]$objectTypeComboBox.Items.Add("(All types)")
        foreach ($typeLabel in ($global:Objects | ForEach-Object { $_.TypeLabel } | Sort-Object -Unique)) {
            [void]$objectTypeComboBox.Items.Add($typeLabel)
        }
        $objectTypeComboBox.SelectedIndex = 0
        $objectPlatformComboBox.SelectedIndex = 0
    } finally {
        $global:SuppressGridEvents = $false
    }

    # An explicit Load or Refresh, and the discovery run above, are deliberate
    # rebuilds - stamp the context to match what was just retrieved.
    $global:DiscoveryContext = "$($global:Mode)|$($global:Action)|$($global:ChangeScope)|$($global:GroupId)|$((@($global:SelectedScopeTags | ForEach-Object { $_.Id }) | Sort-Object) -join ',')"

    Update-ObjectGrid
    Write-LogMessage "$($global:Objects.Count) object(s) listed on the Objects tab."

    if ($global:Objects.Count -eq 0) {
        if ($isRemove -and $byScopeTag) {
            [System.Windows.Forms.MessageBox]::Show("No supported objects carry the selected scope tag(s).", "Nothing To Remove")
        } elseif ($isRemove) {
            [System.Windows.Forms.MessageBox]::Show("No supported objects are assigned to '$($global:SelectedGroup)'.", "Nothing To Remove")
        } else {
            [System.Windows.Forms.MessageBox]::Show("No matching supported objects were found.", "No Results")
        }
    }
}

$loadObjectsButton.Add_Click({ Invoke-ObjectDiscovery })

$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(10, $buttonRowY)
$selectAllButton.Size     = New-Object System.Drawing.Size(110, 35)
$selectAllButton.Text     = "Select All"
$selectAllButton.Font     = $labelFont
$selectAllButton.Add_Click({
    $objectGridView.EndEdit()

    # Deliberately scoped to the rows currently VISIBLE under the filter. That
    # makes "filter to Edge, Select All, clear filter, filter to Defender,
    # Select All" a natural way to build a selection across several filters.
    # Rows hidden by the filter keep whatever state they already had.
    $allChecked = ($objectGridView.Rows.Count -gt 0)
    foreach ($row in $objectGridView.Rows) {
        if ($row.Cells["Select"].Value -ne $true) { $allChecked = $false; break }
    }
    $newValue = -not $allChecked

    # Setting each cell raises CellValueChanged, which keeps the selection set
    # in sync, so the set is not updated directly here.
    foreach ($row in $objectGridView.Rows) {
        $row.Cells["Select"].Value = $newValue
    }
    if ($newValue) { $selectAllButton.Text = "Deselect All" } else { $selectAllButton.Text = "Select All" }
    Update-ObjectCountLabel
})
$tabObjects.Controls.Add($selectAllButton)

$objectCountLabel = New-Object System.Windows.Forms.Label
$objectCountLabel.Location = New-Object System.Drawing.Point(130, ($buttonRowY + 8))
$objectCountLabel.Size     = New-Object System.Drawing.Size(280, 25)
$objectCountLabel.Font     = $labelFont
$tabObjects.Controls.Add($objectCountLabel)

$nextButton2 = New-Object System.Windows.Forms.Button
$nextButton2.Location = New-Object System.Drawing.Point(640, $buttonRowY)
$nextButton2.Size     = New-Object System.Drawing.Size(110, 35)
$nextButton2.Text     = "Next"
$nextButton2.Font     = $labelFont
$nextButton2.Add_Click({
    # Commit any in-progress checkbox edit so the selection set is current.
    $objectGridView.EndEdit()

    # Built from the selection set against the full loaded list, NOT from the
    # visible rows - anything ticked and then filtered out of view must still
    # be included.
    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($item in $global:Objects) {
        if ($global:CheckedObjectIds.Contains([string]$item.Id)) { $selected.Add($item) }
    }

    if ($selected.Count -eq 0) {
        if ($global:Action -eq "Remove") {
            [System.Windows.Forms.MessageBox]::Show("Nothing is selected, so nothing would be removed.", "No Selection")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one object.", "No Selection")
        }
        return
    }
    $global:SelectedObjects = $selected.ToArray()
    $byType = ($global:SelectedObjects | Group-Object TypeLabel | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join ", "
    Write-LogMessage "$($global:SelectedObjects.Count) object(s) selected: $byType"
    Update-SummaryText
    Set-WizardTab -Target $tabSummary -Direction "Next"
})
$tabObjects.Controls.Add($nextButton2)

$backButton2 = New-Object System.Windows.Forms.Button
$backButton2.Location = New-Object System.Drawing.Point(420, $buttonRowY)
$backButton2.Size     = New-Object System.Drawing.Size(100, 35)
$backButton2.Text     = "Back"
$backButton2.Font     = $labelFont
$backButton2.Add_Click({
    if ($global:ChangeScope -eq "Assignments") {
        Set-WizardTab -Target $tabAssignment -Direction "Back"
    } else {
        Set-WizardTab -Target $tabScopeTags -Direction "Back"
    }
})
$tabObjects.Controls.Add($backButton2)

$cancelButton2 = New-Object System.Windows.Forms.Button
$cancelButton2.Location = New-Object System.Drawing.Point(530, $buttonRowY)
$cancelButton2.Size     = New-Object System.Drawing.Size(100, 35)
$cancelButton2.Text     = "Cancel"
$cancelButton2.Font     = $labelFont
$cancelButton2.Add_Click({ $form.Close() })
$tabObjects.Controls.Add($cancelButton2)

# ---------------------------------------------------------------------------
# Tab 3 - Assignment: group search, include/exclude, action, and filter
# ---------------------------------------------------------------------------
$groupSearchLabel = New-Object System.Windows.Forms.Label
$groupSearchLabel.Text     = "Search groups:"
$groupSearchLabel.Location = New-Object System.Drawing.Point(10, 14)
$groupSearchLabel.Size     = New-Object System.Drawing.Size(120, 25)
$groupSearchLabel.Font     = $labelFont
$tabAssignment.Controls.Add($groupSearchLabel)

$groupSearchTextBox = New-Object System.Windows.Forms.TextBox
$groupSearchTextBox.Location = New-Object System.Drawing.Point(135, 12)
$groupSearchTextBox.Size     = New-Object System.Drawing.Size(370, 25)
$groupSearchTextBox.Font     = $labelFont
$groupSearchTextBox.Text     = $defaultGroupSearch
$tabAssignment.Controls.Add($groupSearchTextBox)

$groupSearchButton = New-Object System.Windows.Forms.Button
$groupSearchButton.Location = New-Object System.Drawing.Point(515, 10)
$groupSearchButton.Size     = New-Object System.Drawing.Size(90, 30)
$groupSearchButton.Text     = "Search"
$groupSearchButton.Font     = $labelFont
$tabAssignment.Controls.Add($groupSearchButton)

$groupCountLabel = New-Object System.Windows.Forms.Label
$groupCountLabel.Location = New-Object System.Drawing.Point(615, 17)
$groupCountLabel.Size     = New-Object System.Drawing.Size(135, 25)
$tabAssignment.Controls.Add($groupCountLabel)

$groupGridView = New-Object System.Windows.Forms.DataGridView
$groupGridView.Location           = New-Object System.Drawing.Point(10, 48)
$groupGridView.Size               = New-Object System.Drawing.Size(740, 210)
$groupGridView.AllowUserToAddRows = $false
$groupGridView.RowHeadersVisible  = $false
$groupGridView.ReadOnly           = $true
$groupGridView.MultiSelect        = $false
$groupGridView.SelectionMode      = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

$colGroupName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGroupName.Name       = "DisplayName"
$colGroupName.HeaderText = "Display Name"
$colGroupName.Width      = 470
$groupGridView.Columns.Add($colGroupName) | Out-Null

$colGroupId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGroupId.Name       = "Id"
$colGroupId.HeaderText = "Id"
$colGroupId.Width      = 260
$groupGridView.Columns.Add($colGroupId) | Out-Null
Set-GridHeaderStyle -Grid $groupGridView
$tabAssignment.Controls.Add($groupGridView)

$groupSearchButton.Add_Click({
    $searchText = $groupSearchTextBox.Text
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        [System.Windows.Forms.MessageBox]::Show("Enter at least one character to search for.", "Search Groups")
        return
    }

    try {
        Invoke-WithWaitCursor {
            Write-LogMessage "Searching Entra ID groups for '$searchText'..."
            $global:Groups = @(Search-EntraGroups -SearchText $searchText)
        }
    } catch {
        Write-LogMessage "Error searching groups: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error searching groups. See log: $global:LogPath", "Graph Error")
        return
    }

    Set-GroupGridData -Grid $groupGridView -Items $global:Groups
    $groupCountLabel.Text = "$($global:Groups.Count) match(es)"
    Write-LogMessage "$($global:Groups.Count) group(s) matched '$searchText'."

    if ($global:Groups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No groups matched that search.", "No Results")
    }
})

# Assignment type (Include / Exclude) - own GroupBox so the radios are mutually
# exclusive only with each other.
$typeGroupBox = New-Object System.Windows.Forms.GroupBox
$typeGroupBox.Text     = "Assignment Type"
$typeGroupBox.Location = New-Object System.Drawing.Point(10, 268)
$typeGroupBox.Size     = New-Object System.Drawing.Size(360, 70)
$typeGroupBox.Font     = $labelFont

$includeRadioButton = New-Object System.Windows.Forms.RadioButton
$includeRadioButton.Text     = "Include"
$includeRadioButton.Location = New-Object System.Drawing.Point(20, 28)
$includeRadioButton.Size     = New-Object System.Drawing.Size(120, 25)
$includeRadioButton.Checked  = $true
$typeGroupBox.Controls.Add($includeRadioButton)

$excludeRadioButton = New-Object System.Windows.Forms.RadioButton
$excludeRadioButton.Text     = "Exclude"
$excludeRadioButton.Location = New-Object System.Drawing.Point(160, 28)
$excludeRadioButton.Size     = New-Object System.Drawing.Size(120, 25)
$typeGroupBox.Controls.Add($excludeRadioButton)
$tabAssignment.Controls.Add($typeGroupBox)

# Action (Add / Remove). When Remove is chosen the include/exclude choice and
# the filter are irrelevant (removal matches on group only), so both boxes are
# disabled for clarity.
$actionGroupBox = New-Object System.Windows.Forms.GroupBox
$actionGroupBox.Text     = "Action"
$actionGroupBox.Location = New-Object System.Drawing.Point(390, 268)
$actionGroupBox.Size     = New-Object System.Drawing.Size(360, 70)
$actionGroupBox.Font     = $labelFont

$addRadioButton = New-Object System.Windows.Forms.RadioButton
$addRadioButton.Text     = "Add assignment"
$addRadioButton.Location = New-Object System.Drawing.Point(20, 28)
$addRadioButton.Size     = New-Object System.Drawing.Size(160, 25)
$addRadioButton.Checked  = $true
$actionGroupBox.Controls.Add($addRadioButton)

$removeRadioButton = New-Object System.Windows.Forms.RadioButton
$removeRadioButton.Text     = "Remove assignment"
$removeRadioButton.Location = New-Object System.Drawing.Point(190, 28)
$removeRadioButton.Size     = New-Object System.Drawing.Size(160, 25)
$actionGroupBox.Controls.Add($removeRadioButton)
$tabAssignment.Controls.Add($actionGroupBox)

# Shown only for Remove. The wait is a property of the API, not of this tool -
# worth saying before the operator commits rather than while they are waiting.
$removeScaleLabel = New-Object System.Windows.Forms.Label
$removeScaleLabel.Location  = New-Object System.Drawing.Point(12, 352)
$removeScaleLabel.Size      = New-Object System.Drawing.Size(738, 130)
# Full contrast, not the grey used for incidental hints. This text explains why
# the next step is slow and what it will cost - the operator is meant to read it
# and decide, so it should not look like fine print that can be skipped.
$removeScaleLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$removeScaleLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
$removeScaleLabel.Visible   = $false
$tabAssignment.Controls.Add($removeScaleLabel)

# Assignment filter
$filterGroupBox = New-Object System.Windows.Forms.GroupBox
$filterGroupBox.Text     = "Assignment Filter (Windows devices)"
$filterGroupBox.Location = New-Object System.Drawing.Point(10, 348)
$filterGroupBox.Size     = New-Object System.Drawing.Size(740, 110)
$filterGroupBox.Font     = $labelFont

$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text     = "Filter:"
$filterLabel.Location = New-Object System.Drawing.Point(20, 32)
$filterLabel.Size     = New-Object System.Drawing.Size(70, 25)
$filterGroupBox.Controls.Add($filterLabel)

$filterComboBox = New-Object System.Windows.Forms.ComboBox
$filterComboBox.Location      = New-Object System.Drawing.Point(90, 29)
$filterComboBox.Size          = New-Object System.Drawing.Size(420, 25)
$filterComboBox.Font          = $labelFont
$filterComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$filterGroupBox.Controls.Add($filterComboBox)

$refreshFiltersButton = New-Object System.Windows.Forms.Button
$refreshFiltersButton.Location = New-Object System.Drawing.Point(525, 27)
$refreshFiltersButton.Size     = New-Object System.Drawing.Size(90, 30)
$refreshFiltersButton.Text     = "Refresh"
$refreshFiltersButton.Font     = $labelFont
$refreshFiltersButton.Add_Click({
    try {
        Invoke-WithWaitCursor { $global:AssignmentFilters = @(Get-WindowsAssignmentFilters) }
        $filterComboBox.Items.Clear()
        [void]$filterComboBox.Items.Add("(None)")
        foreach ($assignmentFilter in $global:AssignmentFilters) {
            [void]$filterComboBox.Items.Add($assignmentFilter.DisplayName)
        }
        $filterComboBox.SelectedIndex = 0
        Write-LogMessage "$($global:AssignmentFilters.Count) Windows assignment filter(s) loaded."
    } catch {
        Write-LogMessage "Error loading assignment filters: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error loading assignment filters. See log: $global:LogPath", "Graph Error")
    }
})
$filterGroupBox.Controls.Add($refreshFiltersButton)

$filterModeLabel = New-Object System.Windows.Forms.Label
$filterModeLabel.Text     = "Mode:"
$filterModeLabel.Location = New-Object System.Drawing.Point(20, 74)
$filterModeLabel.Size     = New-Object System.Drawing.Size(70, 25)
$filterGroupBox.Controls.Add($filterModeLabel)

$filterIncludeRadio = New-Object System.Windows.Forms.RadioButton
$filterIncludeRadio.Text     = "Include"
$filterIncludeRadio.Location = New-Object System.Drawing.Point(90, 72)
$filterIncludeRadio.Size     = New-Object System.Drawing.Size(110, 25)
$filterIncludeRadio.Checked  = $true
$filterIncludeRadio.Enabled  = $false
$filterGroupBox.Controls.Add($filterIncludeRadio)

$filterExcludeRadio = New-Object System.Windows.Forms.RadioButton
$filterExcludeRadio.Text     = "Exclude"
$filterExcludeRadio.Location = New-Object System.Drawing.Point(210, 72)
$filterExcludeRadio.Size     = New-Object System.Drawing.Size(110, 25)
$filterExcludeRadio.Enabled  = $false
$filterGroupBox.Controls.Add($filterExcludeRadio)

$filterHintLabel = New-Object System.Windows.Forms.Label
$filterHintLabel.Text     = "Select a filter to enable the mode. Filters already on existing assignments are not modified."
$filterHintLabel.Location = New-Object System.Drawing.Point(340, 74)
$filterHintLabel.Size     = New-Object System.Drawing.Size(380, 35)
$filterGroupBox.Controls.Add($filterHintLabel)
$tabAssignment.Controls.Add($filterGroupBox)

# The filter mode radios only make sense once an actual filter is chosen.
$filterComboBox.Add_SelectedIndexChanged({
    $hasFilter = ($filterComboBox.SelectedIndex -gt 0)
    $filterIncludeRadio.Enabled = $hasFilter
    $filterExcludeRadio.Enabled = $hasFilter
})

# Switch tabs and record the move.
#
# Every step an operator takes is logged, so a session can be reconstructed
# afterwards from the log alone - which matters for a tool that writes to
# production, where "what did they actually click" is the first question when
# something looks wrong.
function Set-WizardTab {
    param (
        [System.Windows.Forms.TabPage]$Target,
        [string]$Direction = "Next"
    )

    $from = "(none)"
    if ($tabControl.SelectedTab) { $from = $tabControl.SelectedTab.Text }
    Write-LogMessage "$Direction : $from -> $($Target.Text)"
    $tabControl.SelectedTab = $Target
}

# Present only the tabs this run needs. A tab that cannot affect the outcome is
# removed rather than disabled: a disabled tab still invites a click and still
# has to be explained.
function Update-TabVisibility {
    $wantAssignment = ($global:ChangeScope -ne "ScopeTags")
    $wantScopeTags  = ($global:ChangeScope -ne "Assignments")

    if ($wantAssignment -and -not $tabControl.TabPages.Contains($tabAssignment)) {
        $tabControl.TabPages.Insert(1, $tabAssignment)
    } elseif (-not $wantAssignment -and $tabControl.TabPages.Contains($tabAssignment)) {
        $tabControl.TabPages.Remove($tabAssignment)
    }

    if ($wantScopeTags -and -not $tabControl.TabPages.Contains($tabScopeTags)) {
        # Directly before Objects, wherever Objects has ended up.
        $tabControl.TabPages.Insert($tabControl.TabPages.IndexOf($tabObjects), $tabScopeTags)
    } elseif (-not $wantScopeTags -and $tabControl.TabPages.Contains($tabScopeTags)) {
        $tabControl.TabPages.Remove($tabScopeTags)
    }
}

# Describe the survey in terms of the mode actually selected. Advising the
# operator to narrow the mode is only useful when there is something to narrow.
function Update-RemoveScaleNote {
    $isRemove = $removeRadioButton.Checked

    switch ($global:Mode) {
        "Apps"     { $scope = "every supported app" }
        "Profiles" { $scope = "every configuration profile, including compliance policies" }
        default    { $scope = "every supported app, configuration profile, and compliance policy" }
    }

    # The backtick on $expand is required - without it PowerShell would expand a
    # variable named $expand into nothing and the sentence would lose its subject.
    $text = "Graph offers no way to ask which objects target a group. Listing them with `$expand=assignments would have answered it in one call, but that is deprecated for apps and returns 503 and 504 errors for configuration profiles. So the wizard reads the assignments of $scope one object at a time, batched twenty to a request. The survey is read-only, but on a large tenant it takes a while and the window will not respond until it finishes."
    if ($global:Mode -eq "Both") {
        $text += [Environment]::NewLine + [Environment]::NewLine +
                 "Both mode surveys apps, configuration profiles, and compliance policies. Choose a narrower mode on the Connect tab to reduce the survey."
    }

    $removeScaleLabel.Text    = $text
    $removeScaleLabel.Visible = $isRemove

    # Filters play no part in a removal - it matches on group alone - so the box
    # is hidden rather than left greyed out underneath the note.
    $filterGroupBox.Visible = -not $isRemove
}

# Remove ignores include/exclude and filters entirely.
$removeRadioButton.Add_CheckedChanged({
    $isRemove = $removeRadioButton.Checked
    $typeGroupBox.Enabled   = -not $isRemove
    $filterGroupBox.Enabled = -not $isRemove
    Update-RemoveScaleNote
})

$nextButton3 = New-Object System.Windows.Forms.Button
$nextButton3.Location = New-Object System.Drawing.Point(640, $buttonRowY)
$nextButton3.Size     = New-Object System.Drawing.Size(110, 35)
$nextButton3.Text     = "Next"
$nextButton3.Font     = $labelFont
$nextButton3.Add_Click({
    if ($groupGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Search for and select a group.", "No Selection")
        return
    }

    $selectedRow          = $groupGridView.SelectedRows[0]
    $global:SelectedGroup = $selectedRow.Cells["DisplayName"].Value
    $global:GroupId       = $selectedRow.Cells["Id"].Value

    if ($removeRadioButton.Checked) { $global:Action = "Remove" } else { $global:Action = "Add" }
    if ($excludeRadioButton.Checked) { $global:AssignmentType = "Exclude" } else { $global:AssignmentType = "Include" }

    # Resolve the filter selection. Index 0 is the "(None)" placeholder, so the
    # backing array index is one less than the combo index.
    if ($global:Action -eq "Add" -and $filterComboBox.SelectedIndex -gt 0) {
        $chosenFilter      = $global:AssignmentFilters[$filterComboBox.SelectedIndex - 1]
        $global:FilterId   = $chosenFilter.Id
        $global:FilterName = $chosenFilter.DisplayName
        if ($filterExcludeRadio.Checked) { $global:FilterMode = "exclude" } else { $global:FilterMode = "include" }
    } else {
        $global:FilterId   = $null
        $global:FilterName = "(None)"
        $global:FilterMode = "none"
    }

    if ($global:Action -eq "Add") {
        Write-LogMessage "Assignment settings: group '$($global:SelectedGroup)' ($($global:GroupId)), action Add, type $($global:AssignmentType), filter '$($global:FilterName)' [$($global:FilterMode)]."
    } else {
        Write-LogMessage "Assignment settings: group '$($global:SelectedGroup)' ($($global:GroupId)), action Remove."
    }

    if ($global:ChangeScope -eq "Assignments") {
        Initialize-ObjectsTab
        Set-WizardTab -Target $tabObjects -Direction "Next"
    } else {
        Initialize-ScopeTagsTab
        Set-WizardTab -Target $tabScopeTags -Direction "Next"
    }
})
$tabAssignment.Controls.Add($nextButton3)

$backButton3 = New-Object System.Windows.Forms.Button
$backButton3.Location = New-Object System.Drawing.Point(420, $buttonRowY)
$backButton3.Size     = New-Object System.Drawing.Size(100, 35)
$backButton3.Text     = "Back"
$backButton3.Font     = $labelFont
$backButton3.Add_Click({ Set-WizardTab -Target $tabConnect -Direction "Back" })
$tabAssignment.Controls.Add($backButton3)

$cancelButton3 = New-Object System.Windows.Forms.Button
$cancelButton3.Location = New-Object System.Drawing.Point(530, $buttonRowY)
$cancelButton3.Size     = New-Object System.Drawing.Size(100, 35)
$cancelButton3.Text     = "Cancel"
$cancelButton3.Font     = $labelFont
$cancelButton3.Add_Click({ $form.Close() })
$tabAssignment.Controls.Add($cancelButton3)

# ---------------------------------------------------------------------------
# Tab 4 - Scope tags (optional, additive only)
# ---------------------------------------------------------------------------
$scopeTagLabel = New-Object System.Windows.Forms.Label
$scopeTagLabel.Text     = "Optionally select scope tags to apply to every selected object:"
$scopeTagLabel.Location = New-Object System.Drawing.Point(10, 10)
$scopeTagLabel.Size     = New-Object System.Drawing.Size(740, 25)
$scopeTagLabel.Font     = $labelFont
$tabScopeTags.Controls.Add($scopeTagLabel)

$scopeTagCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$scopeTagCheckedListBox.Location      = New-Object System.Drawing.Point(10, 45)
$scopeTagCheckedListBox.Size          = New-Object System.Drawing.Size(560, 420)
$scopeTagCheckedListBox.Font          = $labelFont
$scopeTagCheckedListBox.CheckOnClick  = $true
$tabScopeTags.Controls.Add($scopeTagCheckedListBox)

$refreshScopeTagsButton = New-Object System.Windows.Forms.Button
$refreshScopeTagsButton.Location = New-Object System.Drawing.Point(585, 45)
$refreshScopeTagsButton.Size     = New-Object System.Drawing.Size(160, 32)
$refreshScopeTagsButton.Text     = "Refresh Scope Tags"
$refreshScopeTagsButton.Font     = $labelFont
$refreshScopeTagsButton.Add_Click({
    try {
        Invoke-WithWaitCursor { $global:ScopeTags = @(Get-RoleScopeTags) }
        $scopeTagCheckedListBox.Items.Clear()
        foreach ($scopeTag in $global:ScopeTags) {
            [void]$scopeTagCheckedListBox.Items.Add("$($scopeTag.DisplayName)  (ID $($scopeTag.Id))")
        }
        Write-LogMessage "$($global:ScopeTags.Count) scope tag(s) loaded."
    } catch {
        Write-LogMessage "Error loading scope tags: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error loading scope tags. See log: $global:LogPath", "Graph Error")
    }
})
$tabScopeTags.Controls.Add($refreshScopeTagsButton)

# When only scope tags are being changed the Assignment tab is skipped, and with
# it the Add/Remove selector that lives there - so the direction is chosen here
# instead. Shown only in that case, to avoid two controls for one decision.
$scopeTagActionGroupBox = New-Object System.Windows.Forms.GroupBox
$scopeTagActionGroupBox.Text     = "Action"
$scopeTagActionGroupBox.Location = New-Object System.Drawing.Point(585, 130)
$scopeTagActionGroupBox.Size     = New-Object System.Drawing.Size(160, 90)
$scopeTagActionGroupBox.Font     = $labelFont
$scopeTagActionGroupBox.Visible  = $false

$scopeTagAddRadio = New-Object System.Windows.Forms.RadioButton
$scopeTagAddRadio.Text     = "Add tags"
$scopeTagAddRadio.Location = New-Object System.Drawing.Point(15, 26)
$scopeTagAddRadio.Size     = New-Object System.Drawing.Size(130, 24)
$scopeTagAddRadio.Checked  = $true
$scopeTagActionGroupBox.Controls.Add($scopeTagAddRadio)

$scopeTagRemoveRadio = New-Object System.Windows.Forms.RadioButton
$scopeTagRemoveRadio.Text     = "Remove tags"
$scopeTagRemoveRadio.Location = New-Object System.Drawing.Point(15, 54)
$scopeTagRemoveRadio.Size     = New-Object System.Drawing.Size(130, 24)
$scopeTagActionGroupBox.Controls.Add($scopeTagRemoveRadio)
$tabScopeTags.Controls.Add($scopeTagActionGroupBox)

$scopeTagRemoveRadio.Add_CheckedChanged({
    if ($global:ChangeScope -eq "ScopeTags") {
        if ($scopeTagRemoveRadio.Checked) { $global:Action = "Remove" } else { $global:Action = "Add" }
        Initialize-ScopeTagsTab
    }
})

$clearScopeTagsButton = New-Object System.Windows.Forms.Button
$clearScopeTagsButton.Location = New-Object System.Drawing.Point(585, 87)
$clearScopeTagsButton.Size     = New-Object System.Drawing.Size(160, 32)
$clearScopeTagsButton.Text     = "Clear Selection"
$clearScopeTagsButton.Font     = $labelFont
$clearScopeTagsButton.Add_Click({
    for ($i = 0; $i -lt $scopeTagCheckedListBox.Items.Count; $i++) {
        $scopeTagCheckedListBox.SetItemChecked($i, $false)
    }
})
$tabScopeTags.Controls.Add($clearScopeTagsButton)

$scopeTagHintLabel = New-Object System.Windows.Forms.Label
$scopeTagHintLabel.Text     = "Tags not selected here are never touched, and nothing changes if no tags are selected. The built-in Default tag is ID 0."
$scopeTagHintLabel.Location = New-Object System.Drawing.Point(10, 475)
$scopeTagHintLabel.Size     = New-Object System.Drawing.Size(740, 55)
$tabScopeTags.Controls.Add($scopeTagHintLabel)

$nextButton4 = New-Object System.Windows.Forms.Button
$nextButton4.Location = New-Object System.Drawing.Point(640, $buttonRowY)
$nextButton4.Size     = New-Object System.Drawing.Size(110, 35)
$nextButton4.Text     = "Next"
$nextButton4.Font     = $labelFont
$nextButton4.Add_Click({
    $selectedTags = New-Object System.Collections.Generic.List[object]
    foreach ($index in $scopeTagCheckedListBox.CheckedIndices) {
        $selectedTags.Add($global:ScopeTags[$index])
    }
    $global:SelectedScopeTags = $selectedTags.ToArray()

    if ($global:SelectedScopeTags.Count -gt 0) {
        Write-LogMessage "Scope tag selection ($($global:Action)): $((($global:SelectedScopeTags | ForEach-Object { "$($_.DisplayName) [$($_.Id)]" }) -join ', '))."
    } else {
        Write-LogMessage "No scope tags selected; scope tags will not be changed."
    }

    if ($global:ChangeScope -eq "ScopeTags" -and $global:SelectedScopeTags.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one scope tag - this run changes nothing else.", "No Scope Tags Selected")
        return
    }

    Initialize-ObjectsTab
    Set-WizardTab -Target $tabObjects -Direction "Next"
})
$tabScopeTags.Controls.Add($nextButton4)

$backButton4 = New-Object System.Windows.Forms.Button
$backButton4.Location = New-Object System.Drawing.Point(420, $buttonRowY)
$backButton4.Size     = New-Object System.Drawing.Size(100, 35)
$backButton4.Text     = "Back"
$backButton4.Font     = $labelFont
$backButton4.Add_Click({
    if ($global:ChangeScope -eq "ScopeTags") {
        Set-WizardTab -Target $tabConnect -Direction "Back"
    } else {
        Set-WizardTab -Target $tabAssignment -Direction "Back"
    }
})
$tabScopeTags.Controls.Add($backButton4)

$cancelButton4 = New-Object System.Windows.Forms.Button
$cancelButton4.Location = New-Object System.Drawing.Point(530, $buttonRowY)
$cancelButton4.Size     = New-Object System.Drawing.Size(100, 35)
$cancelButton4.Text     = "Cancel"
$cancelButton4.Font     = $labelFont
$cancelButton4.Add_Click({ $form.Close() })
$tabScopeTags.Controls.Add($cancelButton4)

# Compose the summary text. Called from the Objects tab, which is the last tab
# to contribute to it - the object selection is the final input.
function Update-SummaryText {
    if ($global:Action -eq "Remove") {
        $objectLines = ($global:SelectedObjects | ForEach-Object {
            $suffix = ""
            if ($_.PSObject.Properties.Name -contains "AssignmentSummary") { $suffix = "  -> $($_.AssignmentSummary)" }
            "$($_.DisplayName)  [$($_.TypeLabel)]$suffix"
        }) -join "`r`n    "
    } else {
        $objectLines = ($global:SelectedObjects | ForEach-Object { "$($_.DisplayName)  [$($_.TypeLabel)]" }) -join "`r`n    "
    }

    # Removing an exclusion or a sole assignment is easy to do without noticing
    # in a bulk operation, so both are counted out explicitly on the summary.
    $soleCount    = @($global:SelectedObjects | Where-Object { $_.AssignmentSummary -like "*ONLY ASSIGNMENT*" }).Count
    $excludeCount = @($global:SelectedObjects | Where-Object { $_.AssignmentSummary -like "Exclude*" }).Count

    if ($global:SelectedScopeTags.Count -gt 0) {
        $tagLine = ($global:SelectedScopeTags | ForEach-Object { "$($_.DisplayName) (ID $($_.Id))" }) -join ", "
    } else {
        $tagLine = "(no change)"
    }
    if ($global:Action -eq "Remove") { $tagAction = "Scope Tags Rem " } else { $tagAction = "Scope Tags Add " }

    if ($global:Action -eq "Add" -and $global:ChangeScope -ne "ScopeTags") {
        $typeLine   = "Assignment Type : $($global:AssignmentType)`r`n"
        $filterLine = "Filter          : $($global:FilterName) [$($global:FilterMode)]`r`n"
    } else {
        $typeLine   = ""
        $filterLine = ""
    }

    if ($global:ChangeScope -eq "ScopeTags") {
        $groupLines   = ""
        $mechanicLine = "Mechanics       : GET roleScopeTagIds, apply the change, PATCH the result."
    } else {
        $groupLines = "Target Group    : $($global:SelectedGroup)`r`n" +
                      "Group Id        : $($global:GroupId)`r`n"
        if ($global:Mode -eq "Apps") {
            $mechanicLine = "Mechanics       : POST to the app assignments collection (additive by design)."
        } else {
            $mechanicLine = "Mechanics       : GET existing assignments, merge, POST /assign (read-merge-write)."
        }
    }

    $cautionBlock = ""
    if ($global:Action -eq "Remove" -and $global:ChangeScope -eq "ScopeTags") {
        $soleTagCount = @($global:SelectedObjects | Where-Object { $_.AssignmentSummary -like "*ONLY SCOPE TAG*" }).Count
        if ($soleTagCount -gt 0) {
            $cautionBlock = "`r`nWorth checking before Finish:`r`n" +
                            "  $soleTagCount object(s) carry no scope tags other than the ones selected.`r`n" +
                            "  Those objects are skipped rather than left with none - an object with no`r`n" +
                            "  scope tags is not visible to any scoped administrator.`r`n"
        }
    } elseif ($global:Action -eq "Remove" -and ($soleCount -gt 0 -or $excludeCount -gt 0)) {
        $cautionBlock = "`r`nWorth checking before Finish:`r`n"
        if ($soleCount -gt 0) {
            $cautionBlock += "  $soleCount object(s) have this group as their ONLY assignment. Removing it`r`n" +
                             "  leaves them assigned to nobody - still present, but applying to nothing.`r`n"
        }
        if ($excludeCount -gt 0) {
            $cautionBlock += "  $excludeCount object(s) target this group as an EXCLUSION. Removing that`r`n" +
                             "  WIDENS their scope rather than narrowing it.`r`n"
        }
    }

    $summaryTextBox.Text =
        "Mode            : $($global:Mode)`r`n" +
        "Changing        : $($global:ChangeScope)`r`n" +
        "Action          : $($global:Action)`r`n" +
        $typeLine +
        $filterLine +
        $groupLines +
        "$tagAction : $tagLine`r`n" +
        "Tenant Id       : $($global:TenantId)`r`n" +
        "$mechanicLine`r`n" +
        $cautionBlock +
        "`r`n" +
        "Selected Objects ($($global:SelectedObjects.Count)):`r`n    $objectLines"
}

# ---------------------------------------------------------------------------
# Tab 5 - Summary, preview, and execution
# ---------------------------------------------------------------------------
$summaryLabel = New-Object System.Windows.Forms.Label
$summaryLabel.Text     = "Review the actions below. Preview Impact is read-only and is recommended before Finish."
$summaryLabel.Location = New-Object System.Drawing.Point(10, 10)
$summaryLabel.Size     = New-Object System.Drawing.Size(740, 25)
$summaryLabel.Font     = $labelFont
$tabSummary.Controls.Add($summaryLabel)

$summaryTextBox = New-Object System.Windows.Forms.TextBox
$summaryTextBox.Location   = New-Object System.Drawing.Point(10, 40)
$summaryTextBox.Size       = New-Object System.Drawing.Size(740, 495)
$summaryTextBox.Multiline  = $true
$summaryTextBox.ReadOnly   = $true
$summaryTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$summaryTextBox.Font       = $monoFont
$tabSummary.Controls.Add($summaryTextBox)

# Read-only pass that reports what each object looks like today and what the
# result would be. Nothing is written.
$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Location = New-Object System.Drawing.Point(10, $buttonRowY)
$previewButton.Size     = New-Object System.Drawing.Size(140, 35)
$previewButton.Text     = "Preview Impact"
$previewButton.Font     = $labelFont
$previewButton.Add_Click({
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("PREVIEW - no changes have been made.")
    $lines.Add("")

    # Preview reports assignment counts. A scope-tag-only run does not touch
    # assignments, and the Objects tab already lists each object's current tags.
    if ($global:ChangeScope -eq "ScopeTags") {
        $lines.Add("This run changes scope tags only, so there is nothing to preview here.")
        $lines.Add("Each object's current tags are shown on the Objects tab.")
        $summaryTextBox.Text = $lines -join "`r`n"
        return
    }

    try {
        # One batched round trip per 20 objects instead of one GET each.
        $assignmentMap = Invoke-WithWaitCursor { Get-AssignmentMap -Objects $global:SelectedObjects }
    } catch {
        Write-LogMessage "Preview failed: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Preview failed. See log: $global:LogPath", "Graph Error")
        return
    }

    $unreadable = 0
    foreach ($obj in $global:SelectedObjects) {
        $entry = $assignmentMap["$($obj.Id)"]

        # An object whose assignments could not be read must be called out, not
        # shown as having none - "0 assignments" and "could not tell" look
        # identical in a count column and mean opposite things here.
        if ($null -eq $entry -or $entry.Error) {
            $reason = "could not read assignments"
            if ($entry -and $entry.Error) { $reason = $entry.Error }
            $lines.Add(("{0,-45} {1,3}    {2,-3}  UNKNOWN - {3}" -f $obj.DisplayName, "?", "?", $reason))
            Write-LogMessage "Preview could not read assignments for '$($obj.DisplayName)': $reason" "WARN"
            $unreadable++
            continue
        }

        $existing     = @($entry.Value)
        $alreadyThere = @($existing | Where-Object { $_.target.groupId -eq $global:GroupId })

        if ($global:Action -eq "Add") {
            if ($alreadyThere.Count -gt 0) {
                $outcome = "SKIP (group already assigned)"
                $after   = $existing.Count
            } else {
                $outcome = "ADD"
                $after   = $existing.Count + 1
            }
        } else {
            if ($alreadyThere.Count -gt 0) {
                $outcome = "REMOVE $($alreadyThere.Count) assignment(s)"
                $after   = $existing.Count - $alreadyThere.Count
            } else {
                $outcome = "SKIP (group not assigned)"
                $after   = $existing.Count
            }
        }

        $lines.Add(("{0,-45} {1,3} -> {2,-3}  {3}" -f $obj.DisplayName, $existing.Count, $after, $outcome))
    }

    $lines.Add("")
    $lines.Add("Columns: object name, current assignment count, resulting count, outcome.")
    if ($unreadable -gt 0) {
        $lines.Add("")
        $lines.Add("$unreadable object(s) could not be read and are shown as UNKNOWN. Their")
        $lines.Add("assignments were not retrieved, so the effect of Finish on them is not")
        $lines.Add("predicted here. See the log.")
    }
    if ($global:SelectedScopeTags.Count -gt 0) {
        $tagNames = ($global:SelectedScopeTags | ForEach-Object { $_.DisplayName }) -join ", "
        $lines.Add("Scope tags to add (existing tags preserved): $tagNames")
    }

    $summaryTextBox.Text = $lines -join "`r`n"
    Write-LogMessage "Preview completed for $($global:SelectedObjects.Count) object(s)."
})
$tabSummary.Controls.Add($previewButton)

$finishButton = New-Object System.Windows.Forms.Button
$finishButton.Location = New-Object System.Drawing.Point(640, $buttonRowY)
$finishButton.Size     = New-Object System.Drawing.Size(110, 35)
$finishButton.Text     = "Finish"
$finishButton.Font     = $labelFont
$finishButton.Add_Click({
    if ($global:SelectedScopeTags.Count -gt 0) {
        $tagNames = ($global:SelectedScopeTags | ForEach-Object { $_.DisplayName }) -join ', '
        if ($global:Action -eq "Remove") {
            $tagSummary = "`r`nScope tags to remove: $tagNames"
        } else {
            $tagSummary = "`r`nScope tags to add: $tagNames"
        }
    } else {
        $tagSummary = ""
    }

    if ($global:ChangeScope -eq "ScopeTags") {
        $tagNamesOnly = ($global:SelectedScopeTags | ForEach-Object { $_.DisplayName }) -join ', '
        if ($global:Action -eq "Remove") {
            $confirmText = "Remove scope tag(s) $tagNamesOnly from $($global:SelectedObjects.Count) object(s)?" +
                           "`r`n`r`nNo assignments are read or changed. Tags not listed are carried through unchanged, and an object's last remaining tag is never removed."
        } else {
            $confirmText = "Add scope tag(s) $tagNamesOnly to $($global:SelectedObjects.Count) object(s)?" +
                           "`r`n`r`nNo assignments are read or changed. Existing tags are carried through unchanged."
        }
    } elseif ($global:Action -eq "Remove") {
        $confirmText = "Remove group '$($global:SelectedGroup)' from $($global:SelectedObjects.Count) object(s)?$tagSummary" +
                       "`r`n`r`nOnly this group is removed. Every other assignment on those objects is carried through unchanged."
        $soleSelected    = @($global:SelectedObjects | Where-Object { $_.AssignmentSummary -like "*ONLY ASSIGNMENT*" }).Count
        $excludeSelected = @($global:SelectedObjects | Where-Object { $_.AssignmentSummary -like "Exclude*" }).Count
        if ($soleSelected -gt 0) {
            $confirmText += "`r`n`r`n$soleSelected of these have no other assignment and will be left targeting nobody."
        }
        if ($excludeSelected -gt 0) {
            $confirmText += "`r`n`r`n$excludeSelected of these target the group as an EXCLUSION; removing it widens their scope."
        }
    } else {
        $confirmText = "Execute '$($global:Action)' for $($global:SelectedObjects.Count) object(s) against group '$($global:SelectedGroup)'?$tagSummary" +
                       "`r`n`r`nNothing is replaced. Existing assignments and scope tags are carried through unchanged."
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmText, "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $finishButton.Enabled  = $false
    $previewButton.Enabled = $false

    $successCount = 0
    $skipCount    = 0
    $failCount    = 0

    # script: scope so the flag survives being set inside the loop's try/catch.
    $script:sawPermissionError = $false
    $tagUpdated   = 0
    $tagUnchanged = 0
    $tagSkipped   = 0
    $tagFailed    = 0

    $target = New-AssignmentTarget -GroupId $global:GroupId -Type $global:AssignmentType `
                                   -FilterId $global:FilterId -FilterMode $global:FilterMode
    $tagIds = @($global:SelectedScopeTags | ForEach-Object { $_.Id })

    # The wait cursor is set inline rather than via Invoke-WithWaitCursor.
    # That helper invokes its scriptblock with "&", which runs in a CHILD scope,
    # so every counter increment below would write to a throwaway copy and the
    # totals would all report zero. Anything that mutates local variables must
    # run in this scope.
    $previousCursor = [System.Windows.Forms.Cursor]::Current
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        foreach ($obj in $global:SelectedObjects) {
            $objName = $obj.DisplayName

            # --- Assignment ---------------------------------------------
            # Skipped when the run only changes scope tags: there is no group and
            # nothing to merge, so no object's assignments are read or written.
            if ($global:ChangeScope -ne "ScopeTags") {
            try {
                if ($global:Action -eq "Add") {
                    if ($obj.Kind -eq "mobileApps") {
                        $status = Add-AppAssignment -AppId $obj.Id -AppType $obj.OdataType -Target $target
                        switch ($status) {
                            "Added" {
                                Write-LogMessage "Added '$($global:AssignmentType)' (available, filter '$($global:FilterName)'/$($global:FilterMode)) assignment of group '$($global:SelectedGroup)' to app '$objName'."
                                $successCount++
                            }
                            "AlreadyExists" {
                                Write-LogMessage "Group '$($global:SelectedGroup)' is already assigned to app '$objName' as '$($global:AssignmentType)'; skipped." "WARN"
                                $skipCount++
                            }
                            "Conflict" {
                                Write-LogMessage "Group '$($global:SelectedGroup)' is already assigned to app '$objName' with the opposite type; skipped to avoid a conflict." "WARN"
                                $skipCount++
                            }
                        }
                    } else {
                        $result = Add-ProfileAssignment -Kind $obj.Kind -Id $obj.Id -Target $target
                        switch ($result.Status) {
                            "Added" {
                                Write-LogMessage "Added '$($global:AssignmentType)' (filter '$($global:FilterName)'/$($global:FilterMode)) assignment of group '$($global:SelectedGroup)' to profile '$objName' ($($obj.TypeLabel)). Assignments $($result.Before) -> $($result.After)."
                                $successCount++
                            }
                            "AlreadyExists" {
                                Write-LogMessage "Group '$($global:SelectedGroup)' is already assigned to profile '$objName' as '$($global:AssignmentType)'; skipped." "WARN"
                                $skipCount++
                            }
                            "Conflict" {
                                Write-LogMessage "Group '$($global:SelectedGroup)' is already assigned to profile '$objName' with the opposite type; skipped to avoid a conflict." "WARN"
                                $skipCount++
                            }
                        }
                    }
                } else {
                    if ($obj.Kind -eq "mobileApps") {
                        $removed = Remove-AppAssignment -AppId $obj.Id -GroupId $global:GroupId
                        if ($removed -gt 0) {
                            Write-LogMessage "Removed $removed assignment(s) for group '$($global:SelectedGroup)' from app '$objName'."
                            $successCount++
                        } else {
                            Write-LogMessage "No matching assignment for group '$($global:SelectedGroup)' on app '$objName'; nothing to remove." "WARN"
                            $skipCount++
                        }
                    } else {
                        $result = Remove-ProfileAssignment -Kind $obj.Kind -Id $obj.Id -GroupId $global:GroupId
                        if ($result.Status -eq "Removed") {
                            Write-LogMessage "Removed $($result.Removed) assignment(s) for group '$($global:SelectedGroup)' from profile '$objName'. Assignments $($result.Before) -> $($result.After)."
                            $successCount++
                        } else {
                            Write-LogMessage "No matching assignment for group '$($global:SelectedGroup)' on profile '$objName'; nothing to remove." "WARN"
                            $skipCount++
                        }
                    }
                }
            } catch {
                Write-LogMessage "Error processing assignment for '$objName': $_" "ERROR"
                $hint = Get-GraphErrorHint -ErrorRecord $_
                if ($hint) {
                    Write-LogMessage $hint "ERROR"
                    $script:sawPermissionError = $script:sawPermissionError -or ($hint -like "403*")
                }
                $failCount++
            }
            }

            # --- Scope tags (independent of the assignment outcome) ------
            if ($tagIds.Count -gt 0) {
                try {
                    $tagResult = Update-ObjectScopeTag -Kind $obj.Kind -Id $obj.Id -OdataType $obj.OdataType `
                                                       -TagIds $tagIds -Action $global:Action
                    switch ($tagResult.Status) {
                        "Updated" {
                            Write-LogMessage "Scope tags on '$objName': $($tagResult.Before) -> $($tagResult.After)."
                            $tagUpdated++
                        }
                        "SkippedWouldEmpty" {
                            Write-LogMessage "Scope tag removal on '$objName' skipped: it would have left the object with no scope tags." "WARN"
                            $tagSkipped++
                        }
                        default {
                            Write-LogMessage "Scope tags on '$objName' already match the requested change; no change."
                            $tagUnchanged++
                        }
                    }
                } catch {
                    Write-LogMessage "Error updating scope tags on '$objName': $_" "ERROR"
                    $hint = Get-GraphErrorHint -ErrorRecord $_
                    if ($hint) {
                        Write-LogMessage $hint "ERROR"
                        $script:sawPermissionError = $script:sawPermissionError -or ($hint -like "403*")
                    }
                    $tagFailed++
                }
            }
        }
    } finally {
        [System.Windows.Forms.Cursor]::Current = $previousCursor
    }

    Write-LogMessage "Done. Assignments - success: $successCount, skipped: $skipCount, failed: $failCount. Scope tags - updated: $tagUpdated, unchanged: $tagUnchanged, skipped: $tagSkipped, failed: $tagFailed."

    $message = "Operation complete.`r`n`r`n" +
               "Assignments`r`n  Success : $successCount`r`n  Skipped : $skipCount`r`n  Failed  : $failCount`r`n"
    if ($tagIds.Count -gt 0) {
        $message += "`r`nScope Tags ($($global:Action))`r`n  Updated   : $tagUpdated`r`n  Unchanged : $tagUnchanged`r`n  Skipped   : $tagSkipped`r`n  Failed    : $tagFailed`r`n"
        if ($tagSkipped -gt 0) {
            $message += "  (skipped = removal would have left the object with no scope tags)`r`n"
        }
    }
    if ($script:sawPermissionError) {
        $message += "`r`nOne or more operations were refused with 403 Forbidden." +
                    "`r`nThe sign-in succeeded but the identity is not permitted to write." +
                    "`r`nCheck the app registration's consented scopes and the account's" +
                    "`r`nIntune RBAC role. See the log for the full detail.`r`n"
    }
    $message += "`r`nLog: $global:LogPath"

    [System.Windows.Forms.MessageBox]::Show($message, "Finished")
    $form.Close()
})
$tabSummary.Controls.Add($finishButton)

$backButton5 = New-Object System.Windows.Forms.Button
$backButton5.Location = New-Object System.Drawing.Point(420, $buttonRowY)
$backButton5.Size     = New-Object System.Drawing.Size(100, 35)
$backButton5.Text     = "Back"
$backButton5.Font     = $labelFont
$backButton5.Add_Click({ Set-WizardTab -Target $tabObjects -Direction "Back" })
$tabSummary.Controls.Add($backButton5)

$cancelButton5 = New-Object System.Windows.Forms.Button
$cancelButton5.Location = New-Object System.Drawing.Point(530, $buttonRowY)
$cancelButton5.Size     = New-Object System.Drawing.Size(100, 35)
$cancelButton5.Text     = "Cancel"
$cancelButton5.Font     = $labelFont
$cancelButton5.Add_Click({ $form.Close() })
$tabSummary.Controls.Add($cancelButton5)

# ---------------------------------------------------------------------------
# Launch
# ---------------------------------------------------------------------------
Write-LogMessage "Launching Intune Assignment Wizard v11.0. Log file: $global:LogPath"

# Guard against a script function being shadowed by an alias or cmdlet from an
# imported module. PowerShell resolves Alias > Function > Cmdlet, so a collision
# silently redirects calls somewhere else instead of failing - which is exactly
# how Connect-Graph was hijacked by the Microsoft.Graph.Authentication alias.
foreach ($functionName in @("Connect-GraphSession", "Get-GraphCollection", "Get-ProfileAssignment",
                            "Set-ProfileAssignment", "Add-ProfileAssignment", "Remove-ProfileAssignment",
                            "Add-AppAssignment", "Remove-AppAssignment", "Update-ObjectScopeTag")) {
    $resolved = Get-Command $functionName -ErrorAction SilentlyContinue
    if ($resolved -and $resolved.CommandType -ne "Function") {
        Write-LogMessage "Name collision: '$functionName' resolves to a $($resolved.CommandType) from '$($resolved.Source)' instead of this script's function. Rename it." "ERROR"
    }
}
[void]$form.ShowDialog()
