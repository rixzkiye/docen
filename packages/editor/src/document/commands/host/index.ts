import { ClipboardFormatHostCommands, type ClipboardFormatHostView } from "./clipboard-format";
import { CommentsHostCommands, type CommentsHostView } from "./comments";
import { DesignHostCommands, type DesignHostView } from "./design";
import { DialogsOptionsHostCommands, type DialogsOptionsHostView } from "./dialogs-options";
import { DrawingPicturesHostCommands, type DrawingPicturesHostView } from "./drawing-pictures";
import { FieldsHostCommands, type FieldsHostView } from "./fields";
import { FileIoHostCommands, type FileIoHostView } from "./file-io";
import { HeaderFooterHostCommands, type HeaderFooterHostView } from "./header-footer";
import { MailMergeHostCommands, type MailMergeHostView } from "./mail-merge";
import { NavigationViewsHostCommands, type NavigationViewsHostView } from "./navigation-views";
import { ProofingLanguageHostCommands, type ProofingLanguageHostView } from "./proofing-language";
import { ReferencesHostCommands, type ReferencesHostView } from "./references";
import { hostRegistry, type HostCommandRegistry } from "./registry";
import { RevisionsHostCommands, type RevisionsHostView } from "./revisions";
import { SectionsHostCommands, type SectionsHostView } from "./sections-page-setup";
import { TablesHostCommands, type TablesHostView } from "./tables";

export type { HostCommandDomain, HostCommandHandler, HostCommandRegistry } from "./registry";

/**
 * The narrow host views the extracted command domains receive — each domain
 * gets only the accessors its own bodies call (no shared God-object).
 */
export interface HostCommandViews {
  navigation: NavigationViewsHostView;
  sections: SectionsHostView;
  references: ReferencesHostView;
  mailMerge: MailMergeHostView;
  comments: CommentsHostView;
  revisions: RevisionsHostView;
  proofing: ProofingLanguageHostView;
  fields: FieldsHostView;
  clipboard: ClipboardFormatHostView;
  drawing: DrawingPicturesHostView;
  tables: TablesHostView;
  dialogs: DialogsOptionsHostView;
  fileIo: FileIoHostView;
  headerFooter: HeaderFooterHostView;
  design: DesignHostView;
}

/** Assemble the per-domain host commands into the event → handler registry
 *  `#onCommand` consults before the wired Tiptap dispatch. */
export function hostCommands(views: HostCommandViews): HostCommandRegistry {
  return hostRegistry([
    new NavigationViewsHostCommands(views.navigation),
    new SectionsHostCommands(views.sections),
    new ReferencesHostCommands(views.references),
    new MailMergeHostCommands(views.mailMerge),
    new CommentsHostCommands(views.comments),
    new RevisionsHostCommands(views.revisions),
    new ProofingLanguageHostCommands(views.proofing),
    new FieldsHostCommands(views.fields),
    new ClipboardFormatHostCommands(views.clipboard),
    new DrawingPicturesHostCommands(views.drawing),
    new TablesHostCommands(views.tables),
    new DialogsOptionsHostCommands(views.dialogs),
    new FileIoHostCommands(views.fileIo),
    new HeaderFooterHostCommands(views.headerFooter),
    new DesignHostCommands(views.design),
  ]);
}
