using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers the automated validation that runs before any human sees a submission.
    //
    // The strategy document calls this the highest-leverage work in the whole plan, and the reason
    // is arithmetic: the project owner is one volunteer. Every submission rejected automatically with
    // a clear explanation is maintainer time not spent - and it is faster for the contributor too
    // than waiting in a queue to be told the same thing.
    //
    // TWO PROPERTIES ARE TESTED THROUGHOUT, not just the detection:
    //   1. Errors reject and warnings do not, because that line decides what reaches a human.
    //   2. Every finding NAMES ITS SUBJECT, because a contributor with 400 components cannot act
    //      on "a highlight is invalid".
    // ###########################################################################################
    public class SubmissionValidatorTests
    {
        private static SubmissionManifest Valid()
        {
            return new SubmissionManifest
            {
                // Required now, and must agree with Manufacturer/Hardware/Board below - the
                // validator checks both. See SystemDescriptorRules.BuildSystemId.
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Summary = "Corrected the value of R12.",
                Rows = new SubmissionRows
                {
                    Schematics =
                    {
                        new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "main.png" }
                    },
                    Components =
                    {
                        new ComponentEntry { BoardLabel = "U8" },
                        new ComponentEntry { BoardLabel = "R12" }
                    },
                    ComponentHighlights =
                    {
                        new ComponentHighlightEntry
                        {
                            SchematicName = "Main", BoardLabel = "U8",
                            X = "10", Y = "20", Width = "30", Height = "40"
                        }
                    }
                }
            };
        }

        private static readonly string[] ValidFiles = ["main.png"];

        private static IReadOnlyList<ValidationFinding> Validate(
            SubmissionManifest manifest, params string[] files)
        {
            return SubmissionValidator.Validate(
                manifest, files.Length == 0 ? SubmissionValidatorTests.ValidFiles : files);
        }

        // -----------------------------------------------------------------------------------
        // The happy path, and the queue/reject line.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_clean_submission_produces_no_findings_and_can_be_queued()
        {
            IReadOnlyList<ValidationFinding> findings =
                SubmissionValidatorTests.Validate(SubmissionValidatorTests.Valid());

            Assert.Empty(findings);
            Assert.True(SubmissionValidator.CanBeQueued(findings));
        }

        [Fact]
        public void A_submission_with_only_WARNINGS_can_still_be_queued()
        {
            // The line between error and warning decides what reaches a human. A missing summary
            // is unhelpful, not wrong, and blocking on it would reject good data.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Summary = string.Empty;

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.All(findings, finding => Assert.Equal(ValidationSeverity.Warning, finding.Severity));
            Assert.True(SubmissionValidator.CanBeQueued(findings));
        }

        [Fact]
        public void A_single_ERROR_stops_a_submission_being_queued()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });

            Assert.False(SubmissionValidator.CanBeQueued(SubmissionValidatorTests.Validate(manifest)));
        }

        [Fact]
        public void Every_finding_names_what_it_is_about()
        {
            // A contributor with 400 components cannot act on "a highlight is invalid".
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });
            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Nonexistent", BoardLabel = "R12",
                X = "0", Y = "0", Width = "0", Height = "0"
            });

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.NotEmpty(findings);
            Assert.All(findings, finding =>
            {
                Assert.False(string.IsNullOrWhiteSpace(finding.Subject));
                Assert.False(string.IsNullOrWhiteSpace(finding.Message));
                Assert.False(string.IsNullOrWhiteSpace(finding.Code));
            });
        }

        // -----------------------------------------------------------------------------------
        // File references - including the case trap the strategy document names.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_schematic_naming_a_file_that_is_not_in_the_submission_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Schematics[0] = new BoardSchematicEntry
            {
                SchematicName = "Main", SchematicImageFile = "absent.png"
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "file.missing");
        }

        [Fact]
        public void A_file_reference_differing_only_by_CASE_is_an_error_that_explains_itself()
        {
            // THE trap. This works on the contributor's Windows machine and fails on the Linux
            // server, after publication, on somebody else's computer. The message has to say so,
            // or the contributor sees "file not found" for a file they can plainly see.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Schematics[0] = new BoardSchematicEntry
            {
                SchematicName = "Main", SchematicImageFile = "Main.PNG"
            };

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "file.case_mismatch");

            Assert.Equal(ValidationSeverity.Error, finding.Severity);

            // Both spellings must appear, or the fix is not obvious.
            Assert.Contains("Main.PNG", finding.Message, StringComparison.Ordinal);
            Assert.Contains("main.png", finding.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Case_sensitivity_is_NOT_relaxed_to_be_helpful()
        {
            // Pinned deliberately. Making this comparison case-insensitive would make the finding
            // disappear and the data would publish broken. If someone "fixes" it that way, this
            // test is the conversation.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Schematics[0] = new BoardSchematicEntry
            {
                SchematicName = "Main", SchematicImageFile = "MAIN.PNG"
            };

            Assert.False(SubmissionValidator.CanBeQueued(SubmissionValidatorTests.Validate(manifest)));
        }

        [Fact]
        public void A_schematic_with_no_image_file_at_all_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Schematics[0] = new BoardSchematicEntry { SchematicName = "Main" };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "file.unreferenced");
        }

        // -----------------------------------------------------------------------------------
        // Duplicates.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_duplicate_board_label_is_an_error()
        {
            // The board label is the component's natural key, so a duplicate makes row pairing
            // ambiguous and the server cannot diff reliably.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        [Fact]
        public void Board_labels_differing_only_by_case_are_treated_as_duplicates()
        {
            // "U8" and "u8" are the same component to every human who will read this.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "u8" });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        // ###########################################################################################
        // *** REGRESSION (2026-09-23): A COMPONENT-IMAGE ROW WITH ONLY A NOTE IS CORRECT DATA. ***
        //
        // The "Component images" sheet has a Note column beside the File column, and a row may
        // fill in only the Note - "Compatible part-number: BZX55C2V7" on the C64's CR1 and CR2, or
        // Y1's warning that probing the crystal directly can stall the machine.
        //
        // ComponentImageQueries.HasDisplayableImageFile exists precisely so the application can
        // skip these when rendering images, and the published board ships with them. Requiring a
        // file here rejected published data and blocked the whole submission.
        // ###########################################################################################
        [Fact]
        public void A_component_image_row_carrying_only_a_note_is_accepted()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.ComponentImages.Add(new ComponentImageEntry
            {
                BoardLabel = "U8",
                Name = "Pinout",
                File = string.Empty,
                Note = "Compatible part-number: BZX55C2V7",
            });

            Assert.DoesNotContain(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "file.unreferenced");
        }

        // The other half: a row declaring NOTHING is a leftover, not a note, and is still reported.
        // Without this the fix above would simply have stopped checking image files.
        [Fact]
        public void A_component_image_row_with_neither_a_file_nor_a_note_is_still_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.ComponentImages.Add(new ComponentImageEntry
            {
                BoardLabel = "U8",
                Name = "Pinout",
                File = string.Empty,
                Note = string.Empty,
            });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "file.unreferenced");
        }

        // A row that DOES name a file is still checked against the manifest's file list, note or
        // no note - otherwise adding a note would be a way to smuggle a missing image past the
        // validator.
        [Fact]
        public void A_note_does_not_excuse_an_image_naming_a_file_that_is_not_in_the_submission()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.ComponentImages.Add(new ComponentImageEntry
            {
                BoardLabel = "U8",
                Name = "Pinout",
                File = "no-such-image.png",
                Note = "A note that must not hide the missing file.",
            });

            // "file.missing", not "file.unreferenced": the row DOES name a file, that file is just
            // not in the submission. The subject is the file name, which is the actionable part.
            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "file.missing"
                           && finding.Subject.Contains("no-such-image.png", StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // *** REGRESSION (2026-09-23): THE SAME LABEL IN TWO REGIONS IS CORRECT DATA. ***
        //
        // On the C64 250407, C70 is a ceramic capacitor on PAL and a film capacitor on NTSC;
        // U19, Y1, R26, R52 and R53 are region variants too. The application has always supported
        // this - ComponentListBuilder filters components by region, so exactly one of a pair is
        // ever shown - and the PUBLISHED board ships with all six pairs.
        //
        // Keying the duplicate check on the label alone therefore rejected six rows of correct,
        // already-published data and made the board impossible to submit at all. Found on the
        // project owner's first real submission.
        // ###########################################################################################
        [Fact]
        public void The_same_label_in_two_different_regions_is_NOT_a_duplicate()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.Components.Clear();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "PAL" });
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "NTSC" });

            Assert.DoesNotContain(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        // The other half of the same rule: within ONE region the pair really is ambiguous, because
        // the region filter cannot tell the two rows apart. Without this, the fix above would have
        // turned the check off rather than corrected its key.
        [Fact]
        public void The_same_label_TWICE_IN_ONE_region_is_still_a_duplicate()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.Components.Clear();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "PAL" });
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "PAL" });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        // A blank region means "all regions", so two blank-region rows collide with each other -
        // and a blank one does NOT excuse a region-scoped duplicate of itself being missed.
        [Fact]
        public void Two_rows_with_no_region_at_all_are_still_duplicates()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.Components.Clear();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70" });
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "   " });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        // Region matching is case-insensitive for the same reason the label is: nobody reading the
        // board thinks "PAL" and "pal" are different regions.
        [Fact]
        public void Regions_differing_only_by_case_are_the_same_region()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            manifest.Rows.Components.Clear();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "PAL" });
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "C70", Region = "pal" });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.duplicate_label");
        }

        [Fact]
        public void A_component_with_no_board_label_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "   " });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "component.unlabelled");
        }

        [Fact]
        public void A_duplicate_schematic_name_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.Schematics.Add(new BoardSchematicEntry
            {
                SchematicName = "Main", SchematicImageFile = "main.png"
            });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "schematic.duplicate");
        }

        // -----------------------------------------------------------------------------------
        // Highlights.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_highlight_naming_a_schematic_that_does_not_exist_is_an_error()
        {
            // It draws nothing, anywhere - invisible rather than wrong-looking, which is why it
            // has to be caught here rather than noticed later.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Does Not Exist", BoardLabel = "U8",
                X = "1", Y = "1", Width = "1", Height = "1"
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "highlight.unknown_schematic");
        }

        [Theory]
        [InlineData("0", "10")]
        [InlineData("10", "0")]
        [InlineData("-5", "10")]
        public void A_zero_or_negative_sized_highlight_is_an_error(string width, string height)
        {
            // Draws as nothing or as a hairline, and cannot be grabbed and moved - the component
            // looks like it has no highlight and there is no way to fix it from the UI.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Main", BoardLabel = "U8",
                X = "10", Y = "10", Width = width, Height = height
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "highlight.degenerate");
        }

        [Fact]
        public void A_highlight_starting_outside_the_image_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Main", BoardLabel = "U8",
                X = "-50", Y = "10", Width = "10", Height = "10"
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "highlight.negative_origin");
        }

        [Fact]
        public void A_comma_decimal_coordinate_is_REPORTED_not_reinterpreted()
        {
            // The subtlest bug in this file. "1,5" parsed under a comma-decimal culture is 1.5;
            // under an invariant one it is 15 - a highlight ten times too wide, with nothing
            // failing and nobody noticing until they look at the board. The data format is
            // invariant, so a comma is malformed and must be said so.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Main", BoardLabel = "U8",
                X = "10", Y = "10", Width = "1,5", Height = "10"
            };

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "highlight.unparseable");

            // The message must say what the right format is.
            Assert.Contains("dot", finding.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void An_invariant_decimal_coordinate_is_accepted()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Main", BoardLabel = "U8",
                X = "10.5", Y = "20.25", Width = "30.75", Height = "40.125"
            };

            Assert.Empty(SubmissionValidatorTests.Validate(manifest));
        }

        [Theory]
        [InlineData("")]
        [InlineData("abc")]
        [InlineData("1 2")]
        public void An_unparseable_coordinate_is_an_error(string width)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Main", BoardLabel = "U8",
                X = "10", Y = "10", Width = width, Height = "10"
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "highlight.unparseable");
        }

        // -----------------------------------------------------------------------------------
        // Links.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("file:///etc/passwd")]
        [InlineData("javascript:alert(1)")]
        [InlineData("not a url")]
        [InlineData("ftp://example.com/file")]
        public void A_link_that_is_not_http_or_https_is_an_error(string url)
        {
            // These are opened through ExternalTargetLauncher, which admits only http/https/mailto
            // and local paths inside the data root - so this would be refused at click time with a
            // warning nobody sees. Refusing it here is where it can be explained.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Url = url });

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "link.not_web");
        }

        [Theory]
        [InlineData("https://example.com/datasheet.pdf")]
        [InlineData("http://example.com")]
        public void An_ordinary_web_link_is_accepted(string url)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Url = url });

            Assert.Empty(SubmissionValidatorTests.Validate(manifest));
        }

        [Fact]
        public void An_empty_link_is_ignored_rather_than_rejected()
        {
            // A blank cell in a spreadsheet is absence, not an invalid value.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Url = "" });

            Assert.Empty(SubmissionValidatorTests.Validate(manifest));
        }

        // -----------------------------------------------------------------------------------
        // Orphans and scale.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_image_for_a_component_not_in_the_submission_is_a_WARNING_not_an_error()
        {
            // Usually a leftover, but it could be intentional during a staged change, and the data
            // still loads. Rejecting it outright would block a legitimate workflow.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Rows.ComponentImages.Add(new ComponentImageEntry
            {
                BoardLabel = "Q99", File = "main.png"
            });

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "image.orphan");

            Assert.Equal(ValidationSeverity.Warning, finding.Severity);
        }

        [Fact]
        public void An_entirely_empty_submission_is_an_error()
        {
            var manifest = new SubmissionManifest
            {
                Hardware = "C64", Board = "250407", Summary = "Nothing."
            };

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "submission.empty");
        }

        [Fact]
        public void An_implausible_component_count_is_a_warning()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            for (int index = 0; index < SubmissionValidator.ImplausibleComponentCount + 1; index++)
                manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = $"X{index}" });

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "scale.components");

            Assert.Equal(ValidationSeverity.Warning, finding.Severity);
        }

        // -----------------------------------------------------------------------------------
        // Identity and format.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_unknown_format_version_is_an_error_telling_the_contributor_to_update()
        {
            // Silently accepting a payload you cannot fully interpret is how half-applied data
            // gets published.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.FormatVersion = SubmissionFormat.CurrentVersion + 1;

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "format.unsupported");

            Assert.Contains("Update CRT", finding.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void A_submission_that_does_not_name_its_hardware_or_board_is_an_error()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Hardware = string.Empty;
            manifest.Board = string.Empty;

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, finding => finding.Code == "identity.hardware_missing");
            Assert.Contains(findings, finding => finding.Code == "identity.board_missing");
        }

        [Fact]
        public void EVERY_problem_is_reported_in_one_pass()
        {
            // A contributor fixing one problem, resubmitting, and being told the next is a slow
            // loop over a potentially large upload.
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Summary = string.Empty;
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });
            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Ghost", BoardLabel = "R12",
                X = "0", Y = "0", Width = "0", Height = "0"
            });

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, f => f.Code == "summary.missing");
            Assert.Contains(findings, f => f.Code == "component.duplicate_label");
            Assert.Contains(findings, f => f.Code == "highlight.unknown_schematic");
            Assert.Contains(findings, f => f.Code == "highlight.degenerate");
        }

        // -----------------------------------------------------------------------------------
        // The SystemId (Phase 4, task 7).
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // THE ID IS REQUIRED - including for a brand-new system.
        //
        // It is "Manufacturer/Hardware/Board", which the client can always compute from what the
        // contributor typed, so there is no chicken-and-egg where a new system has no identity
        // until the server grants one. What makes a submission NEW is that no `systems` row holds
        // that id yet - a lookup the server does, never a flag the client sets.
        // ###########################################################################################
        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        public void A_submission_with_no_system_id_is_refused(string systemId)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.SystemId = systemId;

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, f => f.Code == "identity.system_id_malformed");
            Assert.False(SubmissionValidator.CanBeQueued(findings));
        }

        // ###########################################################################################
        // A malformed id is REJECTED, because the value is a DATABASE PRIMARY KEY (and a foreign
        // key from `submissions`) and resolves a folder in the published tree. A malformed one is
        // either a broken client or an attempt to attach a submission to something it does not
        // belong to.
        // ###########################################################################################
        [Theory]
        [InlineData("Commodore")]                        // too few segments
        [InlineData("Commodore/C64")]
        [InlineData("Commodore/C64/250407/Data.xlsx")]   // the ExcelDataFile, not the id
        [InlineData("Commodore//250407")]                // empty segment
        [InlineData("Commodore/C64/../etc")]             // traversal
        [InlineData("Commodore/CON/250407")]             // reserved device name
        [InlineData("1; DROP TABLE submissions")]
        public void A_submission_claiming_a_malformed_system_id_is_refused(string systemId)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.SystemId = systemId;

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, f => f.Code == "identity.system_id_malformed");
            Assert.False(SubmissionValidator.CanBeQueued(findings));
        }

        // ###########################################################################################
        // THE ID MUST AGREE WITH THE PARTS IT NAMES.
        //
        // The id is built from Manufacturer/Hardware/Board, so a disagreement means the client
        // assembled the manifest inconsistently - and since everything keys off the id, the parts
        // shown on any screen would silently be the wrong ones. Worth its own finding rather than
        // being folded into "malformed": the shape is fine, the content is not.
        // ###########################################################################################
        [Fact]
        public void A_system_id_disagreeing_with_the_names_it_carries_is_refused()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.SystemId = "Commodore/VIC20/250408";

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, f => f.Code == "identity.system_id_mismatch");
            Assert.False(SubmissionValidator.CanBeQueued(findings));
        }

        // ###########################################################################################
        // *** THE PARTS MUST ALREADY BE CANONICAL (security review, 2026-09-25). ***
        //
        // BuildSystemId trims before joining, so "Commodore" plus a thousand spaces made a VALID,
        // MATCHING id - and the raw part then went into a VARCHAR(100) column and failed the insert
        // with a 500. The client builds the parts from folder names, which are canonical already.
        // ###########################################################################################
        [Theory]
        [InlineData("Commodore ", "C64", "250407")]
        [InlineData("Commodore", "  C64", "250407")]
        [InlineData("Commodore", "C64", "2504  07")]
        public void A_name_part_carrying_extra_spaces_is_refused(string manufacturer, string hardware, string board)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Manufacturer = manufacturer;
            manifest.Hardware = hardware;
            manifest.Board = board;
            manifest.SystemId = SystemDescriptorRules.BuildSystemId(manufacturer, hardware, board);

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "identity.parts_not_canonical");
        }

        // A board sitting WHERE a shared folder lives would make its own folder - which a submission
        // may change freely - the same place as the shared one, turning the scope rules inside out.
        //
        // *** EITHER NAME IN EITHER POSITION (code review, 2026-09-25). *** Only "Generic shared
        // files" as manufacturer and "Shared files" as hardware used to be refused, while
        // DataTreeUsage and PublishedSystemLister skip BOTH names in BOTH positions. A board
        // published as "Shared files/<hw>/<board>" was then never listed for maintainers, and its
        // workbook and every file it cites were reported as unused - and removable.
        [Theory]
        [InlineData("Generic shared files", "C64", "250407")]
        [InlineData("Commodore", "Shared files", "250407")]
        [InlineData("Commodore", "shared FILES", "250407")]
        [InlineData("Shared files", "C64", "250407")]
        [InlineData("Commodore", "Generic shared files", "250407")]
        public void A_board_named_after_a_shared_folder_is_refused(string manufacturer, string hardware, string board)
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Manufacturer = manufacturer;
            manifest.Hardware = hardware;
            manifest.Board = board;
            manifest.SystemId = $"{manufacturer}/{hardware}/{board}";

            Assert.Contains(
                SubmissionValidatorTests.Validate(manifest),
                finding => finding.Code == "identity.reserved_folder");
        }

        // ###########################################################################################
        // *** OVER-LONG FIELDS ARE REFUSED HERE, NOT BY THE DATABASE (security review, 2026-09-25). ***
        //
        // The summary and the revision each land in a bounded column. The revision date reaches its
        // column only AFTER a publish has written the tree - so an over-long one used to turn a
        // successful, irreversible publish into a 500 the maintainer would retry.
        // ###########################################################################################
        [Fact]
        public void An_over_long_summary_is_refused()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Summary = new string('x', SubmissionFormat.MaximumSummaryLength + 1);

            Assert.Contains(SubmissionValidatorTests.Validate(manifest), finding => finding.Code == "summary.too_long");
        }

        [Fact]
        public void An_over_long_revision_or_base_revision_is_refused()
        {
            SubmissionManifest revision = SubmissionValidatorTests.Valid();
            revision.Rows.RevisionDate = new string('x', SubmissionFormat.MaximumRevisionLength + 1);

            SubmissionManifest baseRevision = SubmissionValidatorTests.Valid();
            baseRevision.BaseRevision = new string('x', SubmissionFormat.MaximumRevisionLength + 1);

            Assert.Contains(SubmissionValidatorTests.Validate(revision), finding => finding.Code == "revision.too_long");
            Assert.Contains(SubmissionValidatorTests.Validate(baseRevision), finding => finding.Code == "revision.too_long");
        }

        // Anti-vacuity: exactly at each limit is fine.
        [Fact]
        public void Fields_exactly_at_their_limits_are_accepted()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.Summary = new string('x', SubmissionFormat.MaximumSummaryLength);
            manifest.BaseRevision = new string('x', SubmissionFormat.MaximumRevisionLength);
            manifest.Rows.RevisionDate = new string('x', SubmissionFormat.MaximumRevisionLength);

            Assert.Empty(SubmissionValidatorTests.Validate(manifest));
        }

        [Fact]
        public void An_id_matching_the_names_it_carries_is_accepted()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.DoesNotContain(findings, f => f.Code.StartsWith("identity.system_id", StringComparison.Ordinal));
            Assert.True(SubmissionValidator.CanBeQueued(findings));
        }

        // ###########################################################################################
        // THE REAL FAILURE FROM THE FIRST LIVE SUBMISSION (2026-09-21).
        //
        // The client sent a manifest with blank Hardware and Board but a correct SystemId, because
        // it built the identity from a STALE board entry rather than from the draft's own
        // registration. The server answered with three findings, and this pins the whole set:
        // the two "missing" errors AND the mismatch, which is a CONSEQUENCE of them - an id built
        // from blank parts cannot equal a real one.
        //
        // Worth pinning as a group rather than one finding at a time: it is the combination that
        // identifies this failure, and a future change that made the mismatch check tolerate blank
        // parts would silently turn three clear errors into one baffling one.
        // ###########################################################################################
        [Fact]
        public void A_manifest_with_blank_names_but_a_real_id_reports_all_three_problems()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.SystemId = "Commodore/C64/250407";
            manifest.Hardware = string.Empty;
            manifest.Board = string.Empty;

            IReadOnlyList<ValidationFinding> findings = SubmissionValidatorTests.Validate(manifest);

            Assert.Contains(findings, f => f.Code == "identity.hardware_missing");
            Assert.Contains(findings, f => f.Code == "identity.board_missing");
            Assert.Contains(findings, f => f.Code == "identity.system_id_mismatch");

            Assert.False(SubmissionValidator.CanBeQueued(findings));
        }

        // The message says whose fault it is. A contributor whose data is fine should not go
        // hunting through their board for a problem that is in the software.
        [Fact]
        public void The_malformed_id_message_says_it_is_the_applications_fault_not_the_data()
        {
            SubmissionManifest manifest = SubmissionValidatorTests.Valid();
            manifest.SystemId = "nope";

            ValidationFinding finding = Assert.Single(
                SubmissionValidatorTests.Validate(manifest),
                f => f.Code == "identity.system_id_malformed");

            Assert.Contains("application", finding.Message);
        }
    }
}
