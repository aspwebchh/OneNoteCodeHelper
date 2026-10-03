using System.Threading;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>内置 Agent 与 MCP 共用的无窗口编辑会话。所有草稿操作串行化。</summary>
    internal sealed class AgentEditSession
    {
        private readonly object _gate = new object();
        internal AgentPageSnapshot Snapshot { get; }
        internal AgentTools Tools { get; }
        internal AgentReport Report => Tools.Report;

        internal AgentEditSession(AgentPageSnapshot snapshot, AgentCommitter committer,
            CancellationToken cancellation, AddInSettings settings = null)
        {
            Snapshot = snapshot;
            Tools = new AgentTools(snapshot, committer, cancellation, settings);
        }

        internal object Execute(AgentToolCall call)
        {
            lock (_gate) return Tools.Execute(call);
        }
    }
}
