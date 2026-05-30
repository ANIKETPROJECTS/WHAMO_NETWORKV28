import {
  type InsertProject,
  type UpdateProjectRequest,
  type ProjectResponse
} from "@shared/schema";

export interface IStorage {
  getProjectsByUser(userId: string): Promise<ProjectResponse[]>;
  getProject(id: number): Promise<ProjectResponse | undefined>;
  createProject(project: InsertProject, userId: string): Promise<ProjectResponse>;
  updateProject(id: number, updates: UpdateProjectRequest): Promise<ProjectResponse>;
  deleteProject(id: number): Promise<void>;
}

export class MemStorage implements IStorage {
  private projects: Map<number, ProjectResponse>;
  private currentId: number;

  constructor() {
    this.projects = new Map();
    this.currentId = 1;
  }

  async getProjectsByUser(userId: string): Promise<ProjectResponse[]> {
    return Array.from(this.projects.values())
      .filter(p => p.userId === userId)
      .sort((a, b) => {
        const da = (a.updatedAt || a.createdAt)?.getTime() ?? 0;
        const db = (b.updatedAt || b.createdAt)?.getTime() ?? 0;
        return db - da;
      });
  }

  async getProject(id: number): Promise<ProjectResponse | undefined> {
    return this.projects.get(id);
  }

  async createProject(insertProject: InsertProject, userId: string): Promise<ProjectResponse> {
    const id = this.currentId++;
    const now = new Date();
    const project: ProjectResponse = {
      ...insertProject,
      id,
      userId,
      createdAt: now,
      updatedAt: now,
    };
    this.projects.set(id, project);
    return project;
  }

  async updateProject(id: number, updates: UpdateProjectRequest): Promise<ProjectResponse> {
    const existing = this.projects.get(id);
    if (!existing) throw new Error(`Project ${id} not found`);
    const updated = { ...existing, ...updates, updatedAt: new Date() };
    this.projects.set(id, updated);
    return updated;
  }

  async deleteProject(id: number): Promise<void> {
    this.projects.delete(id);
  }
}

export const storage = new MemStorage();
