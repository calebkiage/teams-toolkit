// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { getLocalizedString } from "../../../common/localizeUtils";
import { DriverContext } from "../interface/commonArgs";
import { TypeSpecCompileArgs } from "./interface/typeSpecCompileArgs";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { err, ok, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import path from "path";
import * as fs from "fs-extra";
import { ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { getParserOptions } from "../../generator/openApiSpec/helper";

const actionName = "typeSpec/compile"; // DO NOT MODIFY the name
const defaultOutputDir = ".generated";
const defaultOpenApiOutputDir = "specs";
const DefaultDAManifestFileName = "declarativeAgent.json";

@Service(actionName) // DO NOT MODIFY the service name
export class TypeSpecCompileDriver implements StepDriver {
  description = getLocalizedString("driver.typeSpec.compile.description");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: TypeSpecCompileArgs,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    const outputs: Map<string, string> = new Map<string, string>();

    const projectPath = ctx.projectPath;
    const mainFilePath = args.path;
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    // const outputDir = path.join(appPackageFolderPath, defaultOutputDir);
    const outputDir = appPackageFolderPath;
    const openApiSpecFolder = path.join(outputDir, defaultOpenApiOutputDir);
    const daManifestFilePath = path.join(outputDir, DefaultDAManifestFileName);
    try {
      if (ctx.ui?.runCommand) {
        // 1. compile type spec to openapi spec
        const tspRes = await ctx.ui.runCommand({
          cmd: `npx --package=@typespec/compiler tsp compile ${mainFilePath} --emit @typespec/openapi3 --emit @microsoft/typespec-copilot-skills --options @microsoft/typespec-copilot-skills.file-type=json --options @microsoft/typespec-copilot-skills.output-file=${DefaultDAManifestFileName} --options @microsoft/typespec-copilot-skills.emitter-output-dir=${outputDir} --options @typespec/openapi3.emitter-output-dir=${openApiSpecFolder}`,
          workingDirectory: ctx.projectPath,
        });

        if (tspRes.isErr()) {
          return {
            result: err(tspRes.error),
            summaries: summaries,
          };
        }

        // 2. call kiota to generate plugin manifest
        const openapiSpecs = await fs.readdir(openApiSpecFolder);
        if (openapiSpecs.length === 0) {
          return {
            result: err(new UserError(actionName, "noOpenApiSpec", "No OpenAPI spec found")),
            summaries: summaries,
          };
        }

        const daManifest = await fs.readJSON(daManifestFilePath);
        const actions = daManifest.actions;
        if (openapiSpecs.length === 1) {
          // only one openapi spec, the spac name should = openapi.yaml
          const spec = openapiSpecs[0];
          if (actions.length !== 1) {
            return {
              result: err(new UserError(actionName, "noOpenApiSpec", "No OpenAPI spec found")),
              summaries: summaries,
            };
          }

          const action = actions[0];
          const pluginManifestName = action.id as string;

          const kiotaRes = await ctx.ui.runCommand({
            cmd: `npx --package=@microsoft/kiota-bundle kiota plugin add -d ${outputDir}/${defaultOpenApiOutputDir}/${spec} --plugin-name ${pluginManifestName} --output ${outputDir}  --type apiplugin`,
            workingDirectory: ctx.projectPath,
            env: {
              KIOTA_CONFIG_PREVIEW: "true",
            },
          });
          if (kiotaRes.isErr()) {
            return {
              result: err(kiotaRes.error),
              summaries: summaries,
            };
          }
        } else {
          for (const spec of openapiSpecs) {
            const action = actions.find((action: any) =>
              spec.toLowerCase().includes((action.id as string).toLowerCase())
            );
            const pluginManifestName = action.id;
            const kiotaRes = await ctx.ui.runCommand({
              cmd: `npx --package=@microsoft/kiota-bundle kiota plugin add -d ${outputDir}/${defaultOpenApiOutputDir}/${spec} --plugin-name ${
                pluginManifestName as string
              } --output ${outputDir}  --type apiplugin`,
              workingDirectory: ctx.projectPath,
              env: {
                KIOTA_CONFIG_PREVIEW: "true",
              },
            });
            if (kiotaRes.isErr()) {
              return {
                result: err(kiotaRes.error),
                summaries: summaries,
              };
            }
          }
        }

        // 2.1 Temp: remove all plugins from Kiota to avoid error when re-provision
        for (const pluginName of actions.map((action: any) => action.id)) {
          const kiotaRemoveRes = await ctx.ui.runCommand({
            cmd: `npx --package=@microsoft/kiota-bundle kiota plugin remove --plugin-name ${
              pluginName as string
            }`,
            workingDirectory: ctx.projectPath,
            env: {
              KIOTA_CONFIG_PREVIEW: "true",
            },
          });

          if (kiotaRemoveRes.isErr()) {
            return {
              result: err(kiotaRemoveRes.error),
              summaries: summaries,
            };
          }
        }

        // 3. Update manifest
        const manifestFilePath = path.join(ctx.projectPath, args.manifestFilePath);
        const manifest = (await fs.readJSON(manifestFilePath)) as TeamsAppManifest;
        if (!manifest.copilotAgents) {
          manifest.copilotAgents = {};
        }

        if (!manifest.copilotAgents.declarativeAgents) {
          manifest.copilotAgents.declarativeAgents = [];
        }

        const relativePath = path
          .relative(path.dirname(manifestFilePath), daManifestFilePath)
          .replace(/\\/g, "/");
        if (
          !manifest.copilotAgents.declarativeAgents.find((agent) => agent.file === relativePath)
        ) {
          manifest.copilotAgents.declarativeAgents.push({
            id: "declarativeAgent",
            file: relativePath,
          });
        }
        await fs.writeJSON(manifestFilePath, manifest, { spaces: 4 });

        return {
          result: ok(outputs),
          summaries: summaries,
        };
      } else {
        return {
          result: err(new UserError(actionName, "noOpenApiSpec", "No OpenAPI spec found")),
          summaries: summaries,
        };
      }
    } catch (error) {
      return {
        result: err(error),
        summaries: summaries,
      };
    }
  }
}
