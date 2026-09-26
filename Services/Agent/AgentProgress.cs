namespace OneNoteCodeHelper.Services.Agent
{
    internal enum AgentStepState { Running, Done, Failed, Note }

    /// <summary>窗口步骤列表里的一行。同一 Id 再报告一次就整行替换（执行中 → 完成或失败）。</summary>
    internal sealed class AgentStep
    {
        internal int Id;
        internal string Text = "";
        internal AgentStepState State;
    }

    /// <summary>
    /// Agent 进度。Turn 为 0、Status 或 Thinking 为 null 表示不变；Thinking 为空串表示收起思考摘录；
    /// Step 不为 null 时新增或更新步骤列表里的一行。思考摘录只在窗口里显示，不写日志。
    /// </summary>
    internal sealed class AgentProgress
    {
        internal int Turn;
        internal string Status;
        internal string Thinking;
        internal AgentStep Step;
    }
}
