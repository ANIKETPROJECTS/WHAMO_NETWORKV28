import { useState } from "react";
import { useProjects, useDeleteProject } from "@/hooks/use-projects";
import { X, FolderOpen, Trash2, Upload, Clock, LayoutGrid, AlertTriangle } from "lucide-react";

interface Project {
  id: number;
  name: string;
  content: any;
  createdAt: string | null;
  updatedAt: string | null;
}

interface ProjectsListPanelProps {
  onClose: () => void;
  onLoadProject: (project: Project) => void;
  onLoadFromFile: () => void;
  currentProjectId: number | null;
}

function formatDate(dateStr: string | null): string {
  if (!dateStr) return "—";
  const d = new Date(dateStr);
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }) +
    " " + d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}

function projectStats(content: any): { nodes: number; pipes: number } {
  if (!content) return { nodes: 0, pipes: 0 };
  const nodes = Array.isArray(content.nodes) ? content.nodes.length : 0;
  const pipes = Array.isArray(content.edges) ? content.edges.length : 0;
  return { nodes, pipes };
}

export function ProjectsListPanel({ onClose, onLoadProject, onLoadFromFile, currentProjectId }: ProjectsListPanelProps) {
  const { data: projects, isLoading, error, refetch } = useProjects();
  const deleteMutation = useDeleteProject();
  const [confirmDeleteId, setConfirmDeleteId] = useState<number | null>(null);

  const handleDelete = async (id: number) => {
    await deleteMutation.mutateAsync(id);
    setConfirmDeleteId(null);
  };

  return (
    <div
      className="fixed inset-0 z-[300] flex items-center justify-center bg-black/50 backdrop-blur-sm"
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl flex flex-col overflow-hidden animate-in zoom-in-95 duration-200"
        style={{ width: "min(860px, 95vw)", height: "min(640px, 90vh)" }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50 rounded-t-2xl flex-shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <FolderOpen className="w-4 h-4 text-white" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900" style={{ fontFamily: "Poppins, sans-serif" }}>
                My Projects
              </h2>
              <p className="text-xs text-slate-500" style={{ fontFamily: "Poppins, sans-serif" }}>
                {Array.isArray(projects) ? `${projects.length} saved project${projects.length !== 1 ? "s" : ""}` : "Loading…"}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={onLoadFromFile}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-medium rounded-lg transition-colors border border-slate-200"
              style={{ fontFamily: "Poppins, sans-serif" }}
              data-testid="btn-load-from-file"
            >
              <Upload className="w-3.5 h-3.5" />
              Import File
            </button>
            <button
              onClick={onClose}
              className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-500 hover:text-slate-700 transition-colors"
              data-testid="btn-projects-close"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-6">
          {isLoading && (
            <div className="flex flex-col items-center justify-center h-full gap-3 text-slate-400">
              <div className="w-8 h-8 border-2 border-slate-200 border-t-blue-500 rounded-full animate-spin" />
              <span className="text-sm" style={{ fontFamily: "Poppins, sans-serif" }}>Loading projects…</span>
            </div>
          )}

          {error && (
            <div className="flex flex-col items-center justify-center h-full gap-3 text-slate-400">
              <AlertTriangle className="w-10 h-10 text-amber-400" />
              <p className="text-sm font-medium text-slate-600" style={{ fontFamily: "Poppins, sans-serif" }}>
                Could not load projects
              </p>
              <button
                onClick={() => refetch()}
                className="text-xs text-blue-600 hover:underline"
                style={{ fontFamily: "Poppins, sans-serif" }}
              >
                Try again
              </button>
            </div>
          )}

          {!isLoading && !error && Array.isArray(projects) && projects.length === 0 && (
            <div className="flex flex-col items-center justify-center h-full gap-4 text-slate-400">
              <div className="w-16 h-16 bg-slate-100 rounded-2xl flex items-center justify-center">
                <LayoutGrid className="w-8 h-8 text-slate-300" />
              </div>
              <div className="text-center">
                <p className="text-sm font-semibold text-slate-500 mb-1" style={{ fontFamily: "Poppins, sans-serif" }}>
                  No saved projects yet
                </p>
                <p className="text-xs text-slate-400" style={{ fontFamily: "Poppins, sans-serif" }}>
                  Click <strong>Save</strong> in the toolbar to save your current network to your account.
                </p>
              </div>
            </div>
          )}

          {!isLoading && !error && Array.isArray(projects) && projects.length > 0 && (
            <div className="flex flex-col gap-2">
              {(projects as Project[]).map((project) => {
                const stats = projectStats(project.content);
                const isActive = project.id === currentProjectId;
                const isDeleting = confirmDeleteId === project.id;

                return (
                  <div
                    key={project.id}
                    className={`group flex items-center gap-4 p-4 rounded-xl border transition-all duration-150 cursor-pointer ${
                      isActive
                        ? "border-blue-300 bg-blue-50/60"
                        : "border-slate-200 hover:border-slate-300 hover:bg-slate-50"
                    }`}
                    data-testid={`project-card-${project.id}`}
                    onClick={() => !isDeleting && onLoadProject(project)}
                  >
                    {/* Icon */}
                    <div className={`w-10 h-10 rounded-lg flex items-center justify-center flex-shrink-0 ${isActive ? "bg-blue-100" : "bg-slate-100 group-hover:bg-slate-200"}`}>
                      <FolderOpen className={`w-5 h-5 ${isActive ? "text-blue-600" : "text-slate-500"}`} />
                    </div>

                    {/* Details */}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <span
                          className="text-sm font-semibold text-slate-800 truncate"
                          style={{ fontFamily: "Poppins, sans-serif" }}
                        >
                          {project.name}
                        </span>
                        {isActive && (
                          <span className="text-[10px] font-bold px-1.5 py-0.5 bg-blue-600 text-white rounded-full flex-shrink-0">
                            OPEN
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-3 mt-0.5">
                        <span className="flex items-center gap-1 text-xs text-slate-400" style={{ fontFamily: "Poppins, sans-serif" }}>
                          <Clock className="w-3 h-3" />
                          {formatDate(project.updatedAt || project.createdAt)}
                        </span>
                        <span className="text-xs text-slate-400" style={{ fontFamily: "Poppins, sans-serif" }}>
                          {stats.nodes} node{stats.nodes !== 1 ? "s" : ""} · {stats.pipes} pipe{stats.pipes !== 1 ? "s" : ""}
                        </span>
                      </div>
                    </div>

                    {/* Actions */}
                    <div className="flex items-center gap-2 flex-shrink-0" onClick={(e) => e.stopPropagation()}>
                      {isDeleting ? (
                        <div className="flex items-center gap-2">
                          <span className="text-xs text-red-600 font-medium" style={{ fontFamily: "Poppins, sans-serif" }}>
                            Delete?
                          </span>
                          <button
                            onClick={() => handleDelete(project.id)}
                            disabled={deleteMutation.isPending}
                            className="px-2.5 py-1 bg-red-600 hover:bg-red-700 text-white text-xs font-semibold rounded-lg transition-colors disabled:opacity-50"
                            style={{ fontFamily: "Poppins, sans-serif" }}
                          >
                            {deleteMutation.isPending ? "…" : "Yes"}
                          </button>
                          <button
                            onClick={() => setConfirmDeleteId(null)}
                            className="px-2.5 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-semibold rounded-lg transition-colors"
                            style={{ fontFamily: "Poppins, sans-serif" }}
                          >
                            No
                          </button>
                        </div>
                      ) : (
                        <>
                          <button
                            onClick={() => onLoadProject(project)}
                            className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                            style={{ fontFamily: "Poppins, sans-serif" }}
                            data-testid={`btn-open-project-${project.id}`}
                          >
                            Open
                          </button>
                          <button
                            onClick={() => setConfirmDeleteId(project.id)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors opacity-0 group-hover:opacity-100"
                            data-testid={`btn-delete-project-${project.id}`}
                            title="Delete project"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
