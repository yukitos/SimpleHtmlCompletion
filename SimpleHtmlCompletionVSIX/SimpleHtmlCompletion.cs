using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Linq;

namespace SimpleHtmlCompletionVSIX
{
    internal class SimpleHtmlCompletionSource : ICompletionSource
    {
        private SimpleHtmlCompletionSourceProvider m_sourceProvider;
        private ITextBuffer m_textBuffer;
        private List<Completion> m_compList;

        public SimpleHtmlCompletionSource(SimpleHtmlCompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            m_sourceProvider = sourceProvider;
            m_textBuffer = textBuffer;
        }

        void ICompletionSource.AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            var tokenSpanPosition = FindTokenSpanAtPosition(session.GetTriggerPoint(m_textBuffer), session);
            var text = tokenSpanPosition.GetText(tokenSpanPosition.TextBuffer.CurrentSnapshot);

            m_compList = new List<Completion>();

            string element = null;
            string attrName = null;
            string attrValue = null;
            if (text.IndexOf('#') > -1)
            {
                var values = text.Split(new char[] { '#' }, 2);
                element = values[0];
                attrName = "id";
                attrValue = values[1];

            }
            else if (text.IndexOf('.') > -1)
            {
                var values = text.Split(new char[] { '.' }, 2);
                element = values[0];
                attrName = "class";
                attrValue = values[1];
            }
            if (!string.IsNullOrWhiteSpace(element))
            {
                m_compList.Add(new Completion(
                    text, string.Format("<{0} {1}=\"{2}\"></{0}>", element, attrName, attrValue), text, null, null));
            }

            var completionSet = new CompletionSet(
                "HTML",
                "HTML",
                tokenSpanPosition,
                m_compList,
                null);
            completionSets.Add(completionSet);
        }

        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            SnapshotPoint currentPoint = (session.TextView.Caret.Position.BufferPosition) - 1;
            ITextStructureNavigator navigator = m_sourceProvider.NavigatorService.GetTextStructureNavigator(m_textBuffer);
            TextExtent firstExtent = navigator.GetExtentOfWord(currentPoint); // text span of id or class
            var secondExtent = navigator.GetExtentOfWord(firstExtent.Span.Start - 1); // delimiter ('#' or '.')
            var thirdExtent = navigator.GetExtentOfWord(secondExtent.Span.Start - 1); // text span of element's name
            var targetSpan = new SnapshotSpan(thirdExtent.Span.Start, firstExtent.Span.End);
            return currentPoint.Snapshot.CreateTrackingSpan(targetSpan, SpanTrackingMode.EdgeInclusive);
        }

        private bool m_isDisposed;
        public void Dispose()
        {
            if (!m_isDisposed)
            {
                GC.SuppressFinalize(this);
                m_isDisposed = true;
            }
        }
    }

    [Export(typeof(ICompletionSourceProvider))]
    [ContentType("htmlx")]
    [Name("html completion")]
    internal class SimpleHtmlCompletionSourceProvider : ICompletionSourceProvider
    {
        [Import]
        internal ITextStructureNavigatorSelectorService NavigatorService { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            return new SimpleHtmlCompletionSource(this, textBuffer);
        }
    }

    [Export(typeof(IVsTextViewCreationListener))]
    [Name("html completion handler")]
    [ContentType("htmlx")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    internal class SimpleHtmlCompletionHandlerProvider : IVsTextViewCreationListener
    {
        [Import]
        internal IVsEditorAdaptersFactoryService AdapterService = null;

        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }

        [Import]
        internal SVsServiceProvider ServiceProvider { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            var textView = AdapterService.GetWpfTextView(textViewAdapter);
            if (textView == null)
                return;

            Func<SimpleHtmlCompletionCommandHandler> createCommandHandler = delegate () { return new SimpleHtmlCompletionCommandHandler(textViewAdapter, textView, this); };
            textView.Properties.GetOrCreateSingletonProperty(createCommandHandler);
        }
    }

    internal class SimpleHtmlCompletionCommandHandler : IOleCommandTarget
    {
        private IOleCommandTarget m_nextCommandHandler;
        private ITextView m_textView;
        private SimpleHtmlCompletionHandlerProvider m_provider;
        private ICompletionSession m_session;

        internal SimpleHtmlCompletionCommandHandler(IVsTextView textViewAdapter, ITextView textView, SimpleHtmlCompletionHandlerProvider provider)
        {
            this.m_textView = textView;
            this.m_provider = provider;

            //add the command to the command chain
            textViewAdapter.AddCommandFilter(this, out m_nextCommandHandler);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (VsShellUtilities.IsInAutomationFunction(m_provider.ServiceProvider))
            {
                return m_nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }

            if (nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB)
            {
                TriggerCompletion();
                var targetCompletionSet = m_session.CompletionSets.FirstOrDefault(i => i.GetType() == typeof(CompletionSet));
                if (targetCompletionSet != null)
                {
                    m_session.SelectedCompletionSet = targetCompletionSet;
                    m_session.SelectedCompletionSet.SelectBestMatch();
                    if (m_session.SelectedCompletionSet.SelectionStatus.IsSelected)
                    {
                        m_session.Commit();
                        return VSConstants.S_OK;
                    }
                }
                else
                {
                    m_session.Dismiss();
                }
            }

            return m_nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        private bool TriggerCompletion()
        {
            //the caret must be in a non-projection location 
            SnapshotPoint? caretPoint =
                m_textView.Caret.Position.Point.GetPoint(
                textBuffer => (!textBuffer.ContentType.IsOfType("projection")), PositionAffinity.Predecessor);
            if (!caretPoint.HasValue)
            {
                return false;
            }

            m_session = m_provider.CompletionBroker.CreateCompletionSession(
                m_textView,
                caretPoint.Value.Snapshot.CreateTrackingPoint(caretPoint.Value.Position, PointTrackingMode.Positive),
                true);

            //subscribe to the Dismissed event on the session 
            m_session.Dismissed += this.OnSessionDismissed;
            m_session.Start();

            return true;
        }

        private void OnSessionDismissed(object sender, EventArgs e)
        {
            m_session.Dismissed -= this.OnSessionDismissed;
            m_session = null;
            Debug.WriteLine("Session ended.");
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return m_nextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }
    }
}
