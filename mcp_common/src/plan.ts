import { z } from "zod";

export const PlanInput = z.object({
  intent: z.enum([
    "create_wall","place_door","place_window",
    "create_floor","create_roof","create_beam",
    "create_column","create_pipe","create_duct"
  ]),
  args: z.record(z.any()).optional(),
});

export type PlanStep = { tool: string; args: any; why: string };
export function orchestratePlan(input: z.infer<typeof PlanInput>): PlanStep[] {
  const { intent, args } = PlanInput.parse(input);
  const steps: PlanStep[] = [];

  // 0) enriquecer contexto antes (si tiene sentido)
  steps.push({ tool: "query_context_snapshot", args: { cacheSec: 30 }, why: "Contexto para validar recursos" });

  // 1) dispatch por intent
  const map: Record<string,string> = {
    create_wall:   "arch_wall_create",
    place_door:    "arch_door_place",
    place_window:  "arch_window_place",
    create_floor:  "arch_floor_create",
    create_roof:   "arch_roof_create",
    create_beam:   "str_beam_create",
    create_column: "str_column_create",
    create_pipe:   "mep_pipe_create",
    create_duct:   "mep_duct_create",
  };
  const tool = map[intent];
  if (tool) steps.push({ tool, args: args ?? {}, why: "Ejecuta la acción principal" });

  // 2) opcional: verificación/QA rápida
  // steps.push({ tool: "qa_fix_pin_all_links", args: {}, why: "Ejemplo de post-step" });

  return steps;
}